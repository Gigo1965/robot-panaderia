"""SaaS de demostracion - Grupo 2 (Inventario), segun el libro de mapas de celdas.

Un modulo por hoja / microservicio, en el orden del Excel maestro:
  Control de Inventario  -> existencia_minima.py (columna G 'Stock', Decisiones Pendientes #63)
                            El resto de la hoja son formulas: se muestra como LECTURA.
  Historico Diario       -> hornadas.py (registrar_hornada, anular_hornada)
  Consumo de Materia Prima -> lo escribe el robot junto con la hornada: no tiene pantalla de captura.

Aqui NO hay reglas de negocio: cada pantalla recoge lo que el mapa marca como "digitado por
la persona" y se lo entrega a la funcion del robot. Si algo no se permite (p. ej. una hornada
sin materia prima, Regla #34), lo rechaza el robot y la pantalla muestra su mensaje.

Permisos (decisiones del usuario, 3 y 4-oct-2026):
  - Control de inventario: lo ven los dos usuarios; la existencia minima la cambia solo administracion.
  - Hornadas: registran y anulan despacho y administracion.
"""
import logging
from datetime import datetime, time as hora_del_dia

from flask import Blueprint, abort, flash, redirect, render_template, request, session, url_for

from . import config                      # noqa: F401  PRIMERO (prepara el entorno del robot)
from . import datos                       # noqa: E402
from .seguridad import con_sesion, envio_repetido, nuevo_envio, solo_administracion     # noqa: E402
from conexion_drive import ConflictoDeVersion, ErrorConexionExcel                       # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla                             # noqa: E402
from conexion_drive.tiempo import ahora_local                                           # noqa: E402
from microservicios.grupo2_inventario import existencia_minima as em                    # noqa: E402
from microservicios.grupo2_inventario import hornadas as hr                             # noqa: E402

log = logging.getLogger("saas_demo")
grupo2 = Blueprint("grupo2", __name__)

HORNADAS_EN_PANTALLA = 25
CANTIDAD_MAXIMA = 9999
RECHAZOS_DEL_ROBOT = (em.MinimoInvalido, hr.HornadaInvalida, LookupError, ValueError)


@grupo2.app_template_filter("cantidad")
def cantidad(valor):
    """6.77 -> '6.77', 10.0 -> '10', 0.7241 -> '0.72'. Las cantidades del inventario, sin ceros de sobra."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return "" if valor in (None, "") else str(valor)
    texto = f"{valor:,.2f}"
    return texto.rstrip("0").rstrip(".") if "." in texto else texto


def _libro_o_problema():
    try:
        return datos.leer_libro(forzar=request.args.get("actualizar") == "1"), None
    except datos.CopiaNoVerificada as error:
        log.error("Copia no verificada: %s", error)
        return None, str(error)
    except (ErrorConexionExcel, KeyError) as error:
        log.exception("No se pudo leer el Excel")
        return None, f"No se pudo leer el Excel de demostración: {error}"


def _avisar(error):
    """Explica en pantalla por que no se escribio."""
    if isinstance(error, RECHAZOS_DEL_ROBOT):
        flash(f"El robot no lo permitió: {error}", "error")
    elif isinstance(error, (datos.Ocupado, datos.CopiaNoVerificada)):
        flash(str(error), "error")
    elif isinstance(error, ConflictoDeVersion):
        flash("El Excel cambió mientras se guardaba y no se pudo terminar. No se escribió nada; inténtalo de nuevo.",
              "error")
    else:
        log.exception("Fallo una escritura del SaaS")
        flash(f"No se pudo guardar en el Excel. No se escribió nada. Detalle: {error}", "error")


# ============================================================================
# Control de Inventario (lectura) y existencia minima, columna G (#63)
# ============================================================================
@grupo2.route("/control-inventario")
@con_sesion
def control_inventario():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        filas = em.leer_existencias(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer Control de Inventario: {error}"), 503
    por_reabastecer = sum(1 for f in filas if f["estado"] == "REABASTECER")
    return render_template("control_inventario.html", libro=libro, filas=filas, por_reabastecer=por_reabastecer,
                           maximo=em.MAXIMO)


@grupo2.route("/control-inventario/minimo", methods=["POST"])
@solo_administracion
def existencia_minima():
    # El producto se entrega al robot tal como esta en la lista (solo sin espacios a los lados):
    # los nombres son LLAVE y se comparan exactos (Regla #16).
    producto = str(request.form.get("producto") or "").strip()[:120]
    minimo = str(request.form.get("minimo") or "").strip()[:20]
    if not producto or not minimo or producto.startswith("="):
        flash("Elige un producto y escribe la existencia mínima. No se escribió nada.", "error")
        return redirect(url_for("grupo2.control_inventario"))
    try:
        resultado, segundos = datos.escribir(
            f"SaaS existencia minima ({session['rol']})",
            lambda: em.actualizar_existencia_minima(producto, minimo, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(url_for("grupo2.control_inventario"))
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"La existencia mínima de {resultado['producto']} ya era {cantidad(resultado['minimo'])}. "
              "No se escribió nada.", "aviso")
    else:
        flash(f"Existencia mínima de {resultado['producto']}: {cantidad(resultado['antes'])} → "
              f"{cantidad(resultado['despues'])} (celda {resultado['celda']}). El estado lo recalculó el Excel. "
              f"Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo2.control_inventario"))


# ============================================================================
# Historico Diario: hornadas
# ============================================================================
def _solo_hora(valor):
    if isinstance(valor, datetime):
        return valor.time()
    return valor if isinstance(valor, hora_del_dia) else None


def _productos_de_hornada(wb):
    """Los productos de produccion propia y la hornada que cada hoja de costo tiene cargada.

    Solo LEE las celdas que hornadas.py declara en HOJAS_COSTO (las mismas del mapa). La hornada
    'cargada' es la ultima registrada de ese producto: es la unica que anular_hornada puede anular.
    """
    productos = []
    for codigo, cfg in hr.HOJAS_COSTO.items():
        hoja = wb[cfg["hoja"]]
        fecha, hora = hoja[cfg["fecha"]].value, _solo_hora(hoja[cfg["hora"]].value)
        unidades = hoja[cfg["cantidad"]].value
        hay = isinstance(fecha, datetime) and isinstance(unidades, (int, float)) and unidades > 0
        productos.append({
            "codigo": codigo, "nombre": str(hoja[cfg["nombre"]].value or codigo).strip(), "hoja": cfg["hoja"],
            "hay_cargada": hay, "fecha": fecha if hay else None, "hora": hora if hay else None,
            "cantidad": unidades if hay else None, "anulada": hoja[cfg["anulada"]].value == 0,
        })
    return productos


def _hornadas(wb):
    ws, tabla = obtener_tabla(wb, "Historico Diario", "HistoricoHorneadas")
    filas = []
    for fila in leer_tabla(ws, tabla):
        unidades = fila.get("Cantidad producida")
        filas.append({
            "fila": fila["_fila"], "fecha": fila.get("Fecha"), "hora": _solo_hora(fila.get("Hora")),
            "producto": fila.get("Producto"), "cantidad": unidades,
            "costo": fila.get("Costo unitario"), "precio": fila.get("Precio unitario"),
            "venta": fila.get("Venta total"), "costo_total": fila.get("Costo total producción"),
            "utilidad": fila.get("Utilidad bruta"),
            "correccion": isinstance(unidades, (int, float)) and unidades < 0,
        })
    filas.sort(key=lambda f: f["fila"], reverse=True)       # lo ultimo que se escribio, arriba
    return filas


@grupo2.route("/hornadas")
@con_sesion
def hornadas():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        productos, lista = _productos_de_hornada(libro["wb"]), _hornadas(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer Historico Diario: {error}"), 503
    return render_template("hornadas.html", libro=libro, productos=productos, hornadas=lista[:HORNADAS_EN_PANTALLA],
                           total_hornadas=len(lista), unico=nuevo_envio(), maximo=CANTIDAD_MAXIMA)


def _codigo_valido(texto):
    codigo = str(texto or "").strip()
    return codigo if codigo in hr.HOJAS_COSTO else None


@grupo2.route("/hornadas/registrar", methods=["POST"])
@con_sesion
def registrar_hornada():
    if envio_repetido(request.form.get("unico")):
        flash("Esa hornada ya se había enviado. No se registró dos veces.", "aviso")
        return redirect(url_for("grupo2.hornadas"))
    codigo = _codigo_valido(request.form.get("producto"))
    try:
        unidades = int(str(request.form.get("cantidad") or "").strip())
    except ValueError:
        unidades = 0
    if codigo is None or not 0 < unidades <= CANTIDAD_MAXIMA:
        flash(f"Elige el producto y una cantidad entera entre 1 y {CANTIDAD_MAXIMA}. No se escribió nada.", "error")
        return redirect(url_for("grupo2.hornadas"))
    momento = ahora_local()          # hora de Costa Rica, aunque el servidor este en otra zona
    try:
        resultado, segundos = datos.escribir(
            f"SaaS hornada {unidades} x {codigo} ({session['rol']})",
            lambda: hr.registrar_hornada(codigo, unidades, momento, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001
        _avisar(error)
        return redirect(url_for("grupo2.hornadas"))
    if resultado.get("estado") == "DUPLICADO":
        flash("Esa hornada ya estaba registrada. No se escribió nada.", "aviso")
    else:
        flash(f"Hornada registrada: {resultado['cantidad']} x {resultado['producto']} (lote {resultado['lote']}). "
              f"Costo unitario ₡{resultado['costo_unitario']:,.2f}. El Excel sumó la producción y rebajó la materia "
              f"prima. Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo2.hornadas"))


@grupo2.route("/hornadas/anular", methods=["POST"])
@con_sesion
def anular_hornada():
    codigo = _codigo_valido(request.form.get("producto"))
    if codigo is None:
        abort(404, "Ese producto no es de producción propia.")
    try:
        resultado, segundos = datos.escribir(
            f"SaaS anular hornada {codigo} ({session['rol']})",
            lambda: hr.anular_hornada(codigo, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001
        _avisar(error)
        return redirect(url_for("grupo2.hornadas"))
    if resultado.get("estado") == "YA_ESTABA_ANULADA":
        flash("Esa hornada ya estaba anulada. No se escribió nada.", "aviso")
    else:
        flash(f"Hornada anulada (lote {resultado['lote']}): se agregó la corrección de {resultado['cantidad_corregida']} "
              f"unidades; la hornada original no se borra. Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo2.hornadas"))
