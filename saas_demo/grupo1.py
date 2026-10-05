"""SaaS de demostracion - Grupo 1 (Captura operativa), segun el libro de mapas de celdas.

Un modulo por hoja / microservicio, en el orden del Excel maestro:
  Ventas despacho        -> comanda_despacho.py  (#70: varios productos, total a cobrar, un guardado)
                            ventas_despacho.py   (anular_venta_despacho; mismas reglas de precio)
  Interface digital      -> pedidos_digitales.py (esta en app.py: entregar y anular)
                            descuento_general.py (celda B5, Decisiones Pendientes #62)
  Detalle comanda digital-> lo escribe el robot junto con el pedido: no tiene pantalla de captura
  Ingreso de Inventario  -> ingreso_inventario.py (registrar_ingresos, sugerencias de nombres)

Aqui NO hay reglas de negocio: cada pantalla recoge lo que el mapa marca como "digitado por
la persona" y se lo entrega a la funcion del robot. Si algo no se permite, lo rechaza el robot.

Permisos (decision del usuario, 3-oct-2026):
  - Ventas despacho: registra y anula el despachador (consulta al gerente antes de anular),
    y tambien administracion.
  - Descuento de dia de pago e Ingreso de Inventario: solo administracion.
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
from microservicios.grupo1_captura_operativa import comanda_despacho as cd              # noqa: E402
from microservicios.grupo1_captura_operativa import descuento_general as dg             # noqa: E402
from microservicios.grupo1_captura_operativa import ingreso_inventario as ii            # noqa: E402
from microservicios.grupo1_captura_operativa import ventas_despacho as vd               # noqa: E402
from microservicios.grupo1_captura_operativa.catalogo import (                          # noqa: E402
    Catalogo, ComboInactivo, PrecioSinCalcular, ProductoNoEncontrado,
)

log = logging.getLogger("saas_demo")
grupo1 = Blueprint("grupo1", __name__)

VENTAS_EN_PANTALLA = 40
INGRESOS_EN_PANTALLA = 25
LINEAS_DE_FACTURA = 8
CANTIDAD_MAXIMA = 9999
RECHAZOS_DEL_ROBOT = (ProductoNoEncontrado, ComboInactivo, PrecioSinCalcular, ii.ProductoNoReconocido,
                      ii.CostoNoDisponible, dg.DescuentoInvalido, cd.TotalCambio, LookupError, ValueError)


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


def _entero(texto):
    try:
        numero = int(str(texto or "").strip())
    except ValueError:
        return None
    return numero if 0 < numero <= CANTIDAD_MAXIMA else None


def _texto_seguro(texto, largo):
    """Texto de una linea, sin '=' al inicio (regla del proyecto: nunca texto que Excel lea como formula)."""
    limpio = " ".join(str(texto or "").split())[:largo]
    return "" if limpio.startswith(("=", "+", "@")) else limpio


# ============================================================================
# Ventas despacho
# ============================================================================
def _ventas(wb):
    ws, tabla = obtener_tabla(wb, vd.HOJA, vd.TABLA)
    filas = []
    for fila in leer_tabla(ws, tabla):
        fecha, hora = fila.get("Fecha"), fila.get("Hora")
        if isinstance(hora, datetime):
            hora = hora.time()
        if not isinstance(fecha, datetime) or not isinstance(hora, hora_del_dia):
            continue
        filas.append({"fila": fila["_fila"], "momento": datetime.combine(fecha.date(), hora),
                      "producto": fila.get("Producto"), "cantidad": fila.get("Cantidad"),
                      "precio": fila.get("Precio"), "total": fila.get("Total"),
                      "anulada": fila.get("Estado de orden") == vd.ORDEN_ANULADA})
    filas.sort(key=lambda f: (f["momento"], f["fila"]), reverse=True)
    return filas


def _catalogo_de_hoy(wb):
    catalogo = Catalogo(wb)
    reglas = catalogo.reglas
    pct = reglas["pct_general"] if reglas["descuento_general"] else 0.0
    combos = {c["codigo"] for c in catalogo.combos}
    productos = [{"codigo": codigo, "nombre": str(nombre).strip(), "precio": round(float(precio) * (1 - pct), 2),
                  "es_combo": codigo in combos}
                 for codigo, nombre, precio in catalogo.lista_para_mostrar()
                 if codigo and isinstance(precio, (int, float))]
    motivo = "feriado" if reglas["feriado"] else ("día de pago" if reglas["dia_pago"] else None)
    return productos, {"motivo": motivo, "pct": reglas["pct_general"] * 100}


def _comanda():
    """La comanda que el despachador esta armando (vive en su sesion, no en el Excel)."""
    return [{"producto": codigo, "cantidad": cantidad} for codigo, cantidad in session.get("comanda", [])]


@grupo1.route("/ventas-despacho")
@con_sesion
def ventas_despacho():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    productos, descuento = _catalogo_de_hoy(libro["wb"])
    ventas = _ventas(libro["wb"])
    cotizacion, problema_comanda = None, None
    if _comanda():
        try:                                    # el total lo calcula el robot, no la pantalla
            cotizacion = cd.cotizar_comanda_despacho(_comanda(), libro["wb"])
        except RECHAZOS_DEL_ROBOT as error:
            problema_comanda = str(error)
    return render_template("ventas_despacho.html", libro=libro, productos=productos, descuento=descuento,
                           ventas=ventas[:VENTAS_EN_PANTALLA], total_ventas=len(ventas), unico=nuevo_envio(),
                           cotizacion=cotizacion, problema_comanda=problema_comanda)


@grupo1.route("/ventas-despacho/agregar", methods=["POST"])
@con_sesion
def agregar_a_comanda():
    """Agrega un producto a la comanda EN PANTALLA. Todavia no se escribe nada en el Excel."""
    libro, problema = _libro_o_problema()
    if libro is None:
        abort(503, problema)
    codigo = _texto_seguro(request.form.get("producto"), 20)
    cantidad = _entero(request.form.get("cantidad"))
    if not codigo or cantidad is None:
        flash(f"Elige un producto y una cantidad entera entre 1 y {CANTIDAD_MAXIMA}.", "error")
        return redirect(url_for("grupo1.ventas_despacho"))
    try:
        cd.cotizar_comanda_despacho(_comanda() + [{"producto": codigo, "cantidad": cantidad}], libro["wb"])
    except RECHAZOS_DEL_ROBOT as error:
        flash(f"El robot no lo permitió: {error}", "error")
        return redirect(url_for("grupo1.ventas_despacho"))
    session["comanda"] = session.get("comanda", []) + [[codigo, cantidad]]
    return redirect(url_for("grupo1.ventas_despacho"))


@grupo1.route("/ventas-despacho/quitar", methods=["POST"])
@con_sesion
def quitar_de_comanda():
    codigo = request.form.get("producto", "")
    if codigo == "*":
        session.pop("comanda", None)
    else:
        session["comanda"] = [l for l in session.get("comanda", []) if l[0].casefold() != codigo.casefold()]
    return redirect(url_for("grupo1.ventas_despacho"))


@grupo1.route("/ventas-despacho/registrar", methods=["POST"])
@con_sesion
def registrar_venta():
    if envio_repetido(request.form.get("unico")):
        flash("Esa venta ya se había enviado. No se registró dos veces.", "aviso")
        return redirect(url_for("grupo1.ventas_despacho"))
    lineas = _comanda()
    if not lineas:
        flash("La comanda está vacía: agrega al menos un producto. No se escribió nada.", "error")
        return redirect(url_for("grupo1.ventas_despacho"))
    try:
        visto = float(request.form.get("total", ""))
    except ValueError:
        abort(400, "El formulario venció. Vuelve a la página anterior, recárgala e inténtalo de nuevo.")
    momento = ahora_local()          # hora de Costa Rica, aunque el servidor este en otra zona
    try:
        resultado, segundos = datos.escribir(
            f"SaaS comanda de despacho, {len(lineas)} linea(s) ({session['rol']})",
            lambda: cd.registrar_comanda_despacho(lineas, momento, total_cotizado=visto,
                                                  origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(url_for("grupo1.ventas_despacho"))
    session.pop("comanda", None)
    if resultado.get("estado") == "DUPLICADO":
        flash("Esa venta ya estaba registrada. No se escribió nada.", "aviso")
    else:
        detalle = ", ".join(f"{l['cantidad']} x {l['producto']}" for l in resultado["lineas"])
        flash(f"Venta registrada: {detalle}. Total a cobrar ₡{resultado['total']:,.2f}. "
              f"Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo1.ventas_despacho"))


@grupo1.route("/ventas-despacho/anular", methods=["POST"])
@con_sesion
def anular_venta():
    libro, problema = _libro_o_problema()
    if libro is None:
        abort(503, problema)
    venta = next((v for v in _ventas(libro["wb"]) if str(v["fila"]) == request.form.get("fila", "")), None)
    if venta is None:
        abort(404, "No existe esa venta en el Excel de demostración.")
    cantidad = venta["cantidad"]
    if isinstance(cantidad, float) and cantidad == int(cantidad):
        cantidad = int(cantidad)
    try:
        resultado, segundos = datos.escribir(
            f"SaaS anular venta despacho fila {venta['fila']} ({session['rol']})",
            lambda: vd.anular_venta_despacho(venta["producto"], cantidad, venta["momento"],
                                             origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001
        _avisar(error)
        return redirect(url_for("grupo1.ventas_despacho"))
    if resultado.get("estado") == "YA_ESTABA_ANULADA":
        flash("Esa venta ya estaba anulada. No se escribió nada.", "aviso")
    else:
        flash(f"Venta anulada: {cantidad} x {venta['producto']}. Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo1.ventas_despacho"))


# ============================================================================
# Interface digital, celda B5: descuento de dia de pago y feriado (#62)
# ============================================================================
@grupo1.route("/descuento-general", methods=["GET", "POST"])
@solo_administracion
def descuento_general():
    if request.method == "POST":
        try:
            resultado, segundos = datos.escribir(
                f"SaaS descuento general ({session['rol']})",
                lambda: dg.actualizar_descuento_general(request.form.get("porcentaje"),
                                                        origen=f"SaaS ({session['rol']})"))
        except Exception as error:          # noqa: BLE001
            _avisar(error)
            return redirect(url_for("grupo1.descuento_general"))
        if resultado.get("estado") == "SIN_CAMBIO":
            flash(f"El descuento ya era {resultado['porcentaje']:g} %. No se escribió nada.", "aviso")
        else:
            flash(f"El descuento quedó en {resultado['porcentaje']:g} %. Guardado en el Excel en {segundos} s.", "ok")
        return redirect(url_for("grupo1.descuento_general"))

    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        actual = dg.leer_descuento_general(libro["wb"])
    except dg.DescuentoInvalido as error:
        return render_template("problema.html", problema=str(error)), 503
    return render_template("descuento_general.html", libro=libro, actual=actual, maximo=dg.MAXIMO_PORCENTAJE)


# ============================================================================
# Ingreso de Inventario (compras)
# ============================================================================
def _productos_de_compra(wb):
    """Lo que se puede COMPRAR: sin produccion propia (entra por Hornadas) y sin codigos de la MP
    de productos que se venden tal como se compran (Regla #32: se compran con su codigo de venta)."""
    return [p for p in ii.leer_productos_inventario(wb)
            if p["movimiento"] == "compra" and not p.get("codigo_venta") and p["codigo"]]


def _ingresos(wb):
    ws, tabla = obtener_tabla(wb, ii.HOJA, ii.TABLA)
    filas = []
    for fila in leer_tabla(ws, tabla):
        compra = fila.get("Compra externa")
        filas.append({
            "fila": fila["_fila"], "fecha": fila.get("Fecha"), "factura": fila.get("Factura"),
            "producto": compra or fila.get("Producción propia"),
            "cantidad": fila.get("Cantidad Compras") if compra else fila.get("Cantidad Producción"),
            "costo": fila.get("Precio unitario") if compra else fila.get("Precio Unitario2"),
            "total": fila.get("Total Compra Externa") if compra else fila.get("Total Producción Propia"),
            "proveedor": fila.get("Proveedor"), "es_compra": bool(compra),
        })
    filas.sort(key=lambda f: f["fila"], reverse=True)
    return filas


def _pantalla_ingreso(libro, valores=None, codigo=200):
    productos = _productos_de_compra(libro["wb"])
    ingresos = _ingresos(libro["wb"])
    proveedores = sorted({str(i["proveedor"]).strip() for i in ingresos
                          if i["es_compra"] and i["proveedor"] not in (None, "")})
    valores = valores or {"fecha": ahora_local().date().isoformat(), "factura": "", "proveedor": "",
                          "lineas": [("", "")] * LINEAS_DE_FACTURA}
    return render_template("ingreso_inventario.html", libro=libro, productos=productos, proveedores=proveedores,
                           ingresos=ingresos[:INGRESOS_EN_PANTALLA], total_ingresos=len(ingresos),
                           valores=valores, unico=nuevo_envio()), codigo


@grupo1.route("/ingreso-inventario", methods=["GET", "POST"])
@solo_administracion
def ingreso_inventario():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    if request.method == "GET":
        return _pantalla_ingreso(libro)

    if envio_repetido(request.form.get("unico")):
        flash("Esa factura ya se había enviado. No se registró dos veces.", "aviso")
        return redirect(url_for("grupo1.ingreso_inventario"))

    productos = request.form.getlist("producto")[:LINEAS_DE_FACTURA]
    cantidades = request.form.getlist("cantidad")[:LINEAS_DE_FACTURA]
    # El producto se entrega al robot tal como se eligio (solo sin espacios a los lados):
    # los nombres son LLAVE y se comparan exactos (Regla #16).
    escrito = [("" if str(p or "").strip().startswith("=") else str(p or "").strip()[:120], str(c or "").strip())
               for p, c in zip(productos, cantidades)]
    escrito += [("", "")] * (LINEAS_DE_FACTURA - len(escrito))
    valores = {"fecha": request.form.get("fecha", "").strip(), "factura": _texto_seguro(request.form.get("factura"), 30),
               "proveedor": _texto_seguro(request.form.get("proveedor"), 60), "lineas": escrito}

    faltas = []
    try:
        fecha = datetime.strptime(valores["fecha"], "%Y-%m-%d")
    except ValueError:
        fecha = None
        faltas.append("la fecha")
    if not valores["factura"]:
        faltas.append("el código de la factura")
    if not valores["proveedor"]:
        faltas.append("el proveedor")
    lineas = []
    for producto, cantidad in escrito:
        if not producto and not cantidad:
            continue
        numero = _entero(cantidad)
        if not producto or numero is None:
            faltas.append(f"una línea incompleta ({producto or 'sin producto'}: cantidad {cantidad or 'vacía'})")
            continue
        lineas.append({"producto": producto, "cantidad": numero})
    if not lineas and not any("línea" in f for f in faltas):
        faltas.append("al menos una línea con producto y cantidad")
    if faltas:
        flash("Falta o está mal: " + "; ".join(faltas) + ". No se escribió nada.", "error")
        return _pantalla_ingreso(libro, valores, 400)

    # La produccion propia NO entra por aqui: la registra la hornada (Grupo 2), que ademas
    # revisa y rebaja la materia prima (Regla #34). Se detiene antes de llamar al robot.
    propios = {p["codigo"].casefold(): p for p in ii.leer_productos_inventario(libro["wb"]) if p["movimiento"] == "produccion"}
    propios.update({str(p["nombre"]).strip().casefold(): p for p in list(propios.values())})
    de_produccion = [l["producto"] for l in lineas if l["producto"].casefold() in propios]
    if de_produccion:
        flash(f"{', '.join(de_produccion)} es producción propia: se registra en Hornadas, no como compra. "
              "No se escribió nada.", "error")
        return _pantalla_ingreso(libro, valores, 400)

    try:
        resultado, segundos = datos.escribir(
            f"SaaS ingreso de inventario {valores['factura']} ({session['rol']})",
            lambda: ii.registrar_ingresos(lineas, valores["factura"], valores["proveedor"], fecha,
                                          origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001
        _avisar(error)
        return _pantalla_ingreso(libro, valores, 400)
    if resultado.get("estado") == "DUPLICADO":
        flash(f"La factura {valores['factura']} ya estaba registrada con esas líneas. No se escribió nada.", "aviso")
    else:
        detalle = "; ".join(f"{r['cantidad']} x {r['producto']} a ₡{r['costo_unitario']:,.2f}"
                            for r in resultado["registradas"])
        repetidas = f" ({len(resultado['duplicadas'])} línea(s) ya estaban y no se repitieron.)" if resultado["duplicadas"] else ""
        flash(f"Factura {valores['factura']} registrada: {detalle}. El costo se tomó de la tabla cruce maestra y "
              f"quedó congelado.{repetidas} Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo1.ingreso_inventario"))
