"""SaaS de demostracion - Grupo 4 (Costeo), segun el libro de mapas de celdas.

Un modulo por hoja, en el orden del Excel maestro:
  costo pan baguette  -> costeo.py, actualizar_precio_referencia (celda G28, Regla #9).
  costo cupcakes      -> costeo.py, actualizar_precio_referencia (celda G27, Regla #9).
      El mapa marca DOS datos digitados en cada hoja: la cantidad de la hornada (G4, con su fecha y
      hora), que ya se registra en la pantalla Hornadas del Grupo 2, y el precio de referencia.
      Todo lo demas son formulas: la receta, los precios de la materia prima (tabla cruce maestra
      MP), los sueldos (planilla) y el agua y la electricidad (gasto operacional). Se muestra como
      LECTURA, tal como lo calculo el Excel.
  planilla, gasto operacional y valor activos -> paquete 2 del Grupo 4, todavia sin construir.

Aqui NO hay reglas de negocio: la pantalla recoge el precio digitado y se lo entrega a la funcion
del robot. Si el precio queda por debajo del costo unitario, lo rechaza el robot y la pantalla
muestra su mensaje.

Permisos (mismo criterio de los Grupos 2 y 3): las lecturas las ven los dos usuarios; el precio de
referencia lo cambia solo administracion (el mapa dice: "lo digita la administracion").
"""
import logging
from datetime import datetime, time as hora_del_dia

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from . import config                      # noqa: F401  PRIMERO (prepara el entorno del robot)
from . import datos                       # noqa: E402
from .seguridad import con_sesion, solo_administracion                     # noqa: E402
from conexion_drive import ConflictoDeVersion, ErrorConexionExcel          # noqa: E402
from microservicios.grupo2_inventario import hornadas as hr                # noqa: E402
from microservicios.grupo4_costeo import costeo as cs                      # noqa: E402

log = logging.getLogger("saas_demo")
grupo4 = Blueprint("grupo4", __name__)

# Codigo del producto -> (titulo de la pantalla, vista). El codigo es el de la tabla cruce maestra.
PANTALLAS_COSTO = {
    "P01": ("Costo pan baguette", "grupo4.costo_baguette"),
    "Pa01": ("Costo cupcakes", "grupo4.costo_cupcakes"),
}
RECHAZOS_DEL_ROBOT = (cs.CosteoInvalido, LookupError, ValueError)


def _colones(valor):
    """430.230228 -> '₡430.23'. Todos los montos del proyecto van con 2 decimales."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return "" if valor in (None, "") else str(valor)
    return f"₡{valor:,.2f}"


def _cantidad(valor):
    """100 -> '100', 0.8 -> '0.8', 0.15 -> '0.15'. Cantidades y factores, sin ceros de sobra."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return "" if valor in (None, "") else str(valor)
    texto = f"{valor:,.2f}"
    return texto.rstrip("0").rstrip(".") if "." in texto else texto


def _porcentaje(valor):
    """0.9587 -> '95.87 %' (la hoja guarda la utilidad como fraccion del costo)."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return ""
    return f"{valor * 100:,.2f} %"


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
# costo pan baguette y costo cupcakes: lectura del costeo y precio de referencia
# ============================================================================
def _hornada_cargada(wb, codigo):
    """La hornada que la hoja de costo tiene cargada (celdas que declara hornadas.py, las del mapa)."""
    cfg = hr.HOJAS_COSTO[codigo]
    hoja = wb[cfg["hoja"]]
    fecha, hora, unidades = hoja[cfg["fecha"]].value, hoja[cfg["hora"]].value, hoja[cfg["cantidad"]].value
    if isinstance(hora, datetime):
        hora = hora.time()
    hay = isinstance(fecha, datetime) and isinstance(unidades, (int, float)) and unidades > 0
    return {"hay": hay, "fecha": fecha if hay else None,
            "hora": hora if hay and isinstance(hora, hora_del_dia) else None,
            "cantidad": unidades if hay else None, "celda_cantidad": cfg["cantidad"],
            "anulada": hoja[cfg["anulada"]].value == 0}


def _pantalla_costo(codigo):
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    titulo = PANTALLAS_COSTO[codigo][0]
    try:
        producto = next(p for p in cs.leer_costeo(libro["wb"]) if p["codigo"] == codigo)
        hornada = _hornada_cargada(libro["wb"], codigo)
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer la hoja de {titulo.lower()}: {error}"), 503

    for seccion in producto["secciones"]:
        seccion["total_texto"] = _colones(seccion["total"])
        for fila in seccion["filas"]:
            fila.update(cantidad_texto=_cantidad(fila["cantidad"]), precio_texto=_colones(fila["precio"]),
                        costo_texto=_colones(fila["costo"]))
    for renglon in producto["resumen"]:
        es_unidades = "unidades" in renglon["etiqueta"].lower()
        renglon["texto"] = _cantidad(renglon["valor"]) if es_unidades else _colones(renglon["valor"])
    cifras = {dato: _colones(producto[dato]) for dato in
              ("costo_unitario", "precio_referencia", "calculado", "utilidad", "utilidad_calculada")}
    cifras["porcentaje"] = _porcentaje(producto["porcentaje"])
    cifras["porcentaje_calculado"] = _porcentaje(producto["porcentaje_calculado"])
    otras = [(texto, vista) for otro, (texto, vista) in PANTALLAS_COSTO.items() if otro != codigo]
    return render_template("costo.html", libro=libro, titulo=titulo, p=producto, hornada=hornada, cifras=cifras,
                           otras=otras)


@grupo4.route("/costo-pan-baguette")
@con_sesion
def costo_baguette():
    return _pantalla_costo("P01")


@grupo4.route("/costo-cupcakes")
@con_sesion
def costo_cupcakes():
    return _pantalla_costo("Pa01")


@grupo4.route("/costo/precio-referencia", methods=["POST"])
@solo_administracion
def precio_referencia():
    codigo = str(request.form.get("producto") or "").strip()[:10]
    precio = str(request.form.get("precio") or "").strip()[:20]
    destino = url_for(PANTALLAS_COSTO.get(codigo, PANTALLAS_COSTO["P01"])[1])
    if not codigo or not precio or precio.startswith("="):
        flash("Escribe el precio de referencia antes de guardar. No se escribió nada.", "error")
        return redirect(destino)
    try:
        resultado, segundos = datos.escribir(
            f"SaaS precio de referencia {codigo} ({session['rol']})",
            lambda: cs.actualizar_precio_referencia(codigo, precio, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(destino)
    destino = url_for(PANTALLAS_COSTO[resultado["codigo"]][1])
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"El precio de referencia de {resultado['producto']} ya era {_colones(resultado['precio'])}. "
              "No se escribió nada.", "aviso")
    else:
        cambio = resultado["cambios"][0]
        flash(f"Precio de referencia de {resultado['producto']}: {_colones(cambio['antes'])} → "
              f"{_colones(cambio['despues'])} (hoja {resultado['hoja']}, celda {cambio['celda']}). El Excel recalculó "
              f"la utilidad, la tabla cruce maestra y los combos que llevan {resultado['producto']}. "
              f"Guardado en el Excel en {segundos} s.", "ok")
    for aviso in resultado.get("avisos", []):
        flash(aviso, "aviso")
    return redirect(destino)
