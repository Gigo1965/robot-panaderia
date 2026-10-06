"""SaaS de demostracion - Grupo 5 (Financiero), segun el libro de mapas de celdas.

Un modulo por hoja, en el orden del Excel maestro. Paquete 1 (6-oct-2026):
  Flujo caja diario -> finanzas.py, registrar_caja_chica (columna G, Decisiones Pendientes #66).
      El mapa: "G: Input digitado por el usuario final EN EL SaaS -> el ROBOT escribe aqui. UNICA
      columna donde el robot ESCRIBE en esta hoja". Todo lo demas (ventas, compras, egresos, saldo)
      son formulas y se muestra como LECTURA, igual que el saldo de caja actual (N13), que el mapa
      llama "celda de salida principal hacia el SaaS".
  Flujo caja anual  -> finanzas.py, registrar_salida_efectivo (B19:M19, Decisiones Pendientes #67).
      El mapa: "La UNICA EXCEPCION es 'Salidas de efectivo' (B19:M19): una persona lo digita en la
      interfaz del SaaS y el robot lo escribe como VALOR FIJO en la celda del mes".
Pendiente (paquete 2): flujo caja proyecto, Punto de Equilibrio, TIR VAR y las dos proyecciones.

Aqui NO hay reglas de negocio: la pantalla recoge lo digitado y se lo entrega a la funcion del robot.

Permisos (DECISION DEL USUARIO, 6-oct-2026): todo el Grupo 5 es SOLO para administracion, igual que
las hojas de costo (el servidor responde 403 a despacho y su menu no ofrece estas pantallas).
Caja chica (misma decision): se digita el TOTAL del dia con su motivo y reemplaza al anterior.
"""
import logging

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from . import config                      # noqa: F401  PRIMERO (prepara el entorno del robot)
from . import datos                       # noqa: E402
from .seguridad import solo_administracion                                 # noqa: E402
from conexion_drive import ConflictoDeVersion, ErrorConexionExcel          # noqa: E402
from microservicios.grupo5_financiero import finanzas as fz                # noqa: E402

log = logging.getLogger("saas_demo")
grupo5 = Blueprint("grupo5", __name__)

RECHAZOS_DEL_ROBOT = (fz.FinanzasInvalido, LookupError, ValueError)
MESES_PANTALLA = ("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")
MESES_FECHA = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")


def _colones(valor):
    """3500 -> '₡3,500.00'. Todos los montos del proyecto van con 2 decimales."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return ""
    return f"₡{valor:,.2f}"


def _cantidad(valor):
    """1400 -> '1,400'. Unidades, sin ceros de sobra."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return ""
    texto = f"{valor:,.2f}"
    return texto.rstrip("0").rstrip(".") if "." in texto else texto


def _dia(fecha):
    return f"{fecha.day} {MESES_FECHA[fecha.month - 1]} {fecha.year}"


def _dato(nombre, largo=40):
    return str(request.form.get(nombre) or "").strip()[:largo]


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
# Flujo caja diario: lectura y caja chica del dia (columna G)
# ============================================================================
@grupo5.route("/flujo-caja-diario")
@solo_administracion
def flujo_diario():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        leido = fz.leer_flujo_diario(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer la hoja Flujo caja diario: {error}"), 503
    for dia in leido["dias"]:
        dia["fecha_texto"] = _dia(dia["fecha"])
        dia.update({f"{dato}_texto": _colones(dia[dato]) for dato in fz.COLUMNAS_DIARIO.values()})
    for renglon in leido["resumen"]:
        renglon["total_texto"] = _colones(renglon["total"])
    sin_calcular = [d for d in leido["dias"] if not d["calculado"]]
    return render_template("flujo_caja_diario.html", libro=libro, leido=leido, saldo_texto=_colones(leido["saldo"]),
                           hoy_texto=_dia(leido["hoy"]), primera_texto=_dia(leido["primera"]),
                           sin_calcular=len(sin_calcular), tope_texto=_colones(leido["tope"]),
                           largo_motivo=fz.LARGO_MOTIVO)


@grupo5.route("/flujo-caja-diario/caja-chica", methods=["POST"])
@solo_administracion
def caja_chica():
    fecha, monto = _dato("fecha", 12), _dato("monto", 15)
    motivo = " ".join(str(request.form.get("motivo") or "").split())[:fz.LARGO_MOTIVO + 1]
    destino = url_for("grupo5.flujo_diario")
    if not fecha or not monto or not motivo or monto.startswith("="):
        flash("Elige el día y escribe el total y el motivo. No se escribió nada.", "error")
        return redirect(destino)
    try:
        resultado, segundos = datos.escribir(
            f"SaaS caja chica {fecha} ({session['rol']})",
            lambda: fz.registrar_caja_chica(fecha, monto, motivo, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(destino)
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"La caja chica del {_dia(resultado['fecha'])} ya era {_colones(resultado['despues'])}. "
              "No se escribió nada.", "aviso")
    else:
        flash(f"Caja chica del {_dia(resultado['fecha'])}: {_colones(resultado['antes'])} → "
              f"{_colones(resultado['despues'])} (hoja {resultado['hoja']}, celda {resultado['celda']}). "
              f"Motivo: {resultado['motivo']}. El Excel recalculó los egresos de ese día y el saldo de caja. "
              f"Guardado en el Excel en {segundos} s.", "ok")
    return redirect(destino)


# ============================================================================
# Flujo caja anual: lectura y salidas de efectivo del mes (fila 19)
# ============================================================================
@grupo5.route("/flujo-caja-anual")
@solo_administracion
def flujo_anual():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        leido = fz.leer_flujo_anual(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer la hoja flujo caja anual: {error}"), 503
    for fila in leido["filas"]:
        como = _cantidad if "cantidad" in fila["etiqueta"].casefold() else _colones
        fila.update(meses_texto=[como(valor) for valor in fila["meses"]], total_texto=como(fila["total"]))
    meses = [(numero, corto, largo.capitalize())
             for numero, (corto, largo) in enumerate(zip(MESES_PANTALLA, leido["meses"]), 1)]
    return render_template("flujo_caja_anual.html", libro=libro, leido=leido, meses=meses,
                           formulas=sum(1 for f in leido["filas"] if f["clase"] == "formula"))


@grupo5.route("/flujo-caja-anual/salidas", methods=["POST"])
@solo_administracion
def salidas_efectivo():
    mes, monto = _dato("mes", 12), _dato("monto", 15)
    destino = url_for("grupo5.flujo_anual")
    if not mes or not monto or monto.startswith("="):
        flash("Elige el mes y escribe el monto. No se escribió nada.", "error")
        return redirect(destino)
    try:
        resultado, segundos = datos.escribir(
            f"SaaS salidas de efectivo mes {mes} ({session['rol']})",
            lambda: fz.registrar_salida_efectivo(mes, monto, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001
        _avisar(error)
        return redirect(destino)
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"Las salidas de efectivo de {resultado['nombre_mes']} ya eran {_colones(resultado['despues'])}. "
              "No se escribió nada.", "aviso")
    else:
        flash(f"Salidas de efectivo de {resultado['nombre_mes']}: {_colones(resultado['antes'])} → "
              f"{_colones(resultado['despues'])} (hoja {resultado['hoja']}, celda {resultado['celda']}). "
              f"Queda como valor fijo de ese mes. Guardado en el Excel en {segundos} s.", "ok")
    return redirect(destino)
