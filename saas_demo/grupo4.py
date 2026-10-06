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
      || 5-oct-2026 (noche), paquete 2:
  costo pan baguette / costo cupcakes -> actualizar_proporcion: la parte del sueldo de panadero y ayudante
      y la parte del agua y de la electricidad que se carga a cada producto (DECISION DEL USUARIO: las
      define la administracion). Segundo formulario en la misma pantalla.
  planilla          -> actualizar_plaza (personas y sueldo, columnas B y C) y agregar_plaza (filas 13 a 32).
  gasto operacional -> actualizar_gasto (los 9 rubros manuales, un monto por mes).
  valor activos     -> agregar_activo y actualizar_activo (dar de baja = cantidad 0).

Aqui NO hay reglas de negocio: la pantalla recoge el precio digitado y se lo entrega a la funcion
del robot. Si el precio queda por debajo del costo unitario, lo rechaza el robot y la pantalla
muestra su mensaje.

Permisos (mismo criterio de los Grupos 2 y 3): las lecturas las ven los dos usuarios; el precio de
referencia lo cambia solo administracion (el mapa dice: "lo digita la administracion").
Paquete 2: igual. Las tres hojas nuevas las ven los dos usuarios (decision del 3-oct-2026: despacho
"puede ver las lecturas") y las cambia solo administracion.
|| 5-oct-2026 (noche), DECISION DEL USUARIO: "dejar solo a administracion ver las planillas y las hojas de
costos". Las pantallas Costo pan baguette, Costo cupcakes y Planilla son SOLO para administracion (el
servidor responde 403 a despacho y el menu no las ofrece). Gasto operacional y Valor activos siguen como
lectura para los dos usuarios.
|| Mismo dia, DECISION DEL USUARIO: "Gasto operacional no visible para despacho" (trae la fila Planilla y
las cargas sociales). Tambien es SOLO para administracion. Valor activos sigue como lectura para los dos.
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
MESES_PANTALLA = ("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")


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


def _por_ciento(fraccion):
    """0.8 -> '80 %', 0.125 -> '12.5 %'. La celda guarda la fraccion."""
    if isinstance(fraccion, bool) or not isinstance(fraccion, (int, float)):
        return ""
    return f"{round(fraccion * 100, 2):g} %"


def _dato(texto, largo=40):
    """Un campo del formulario; vacio -> None (el robot entiende 'no cambia')."""
    valor = str(request.form.get(texto) or "").strip()[:largo]
    return valor or None


def _cambios(resultado, dinero=()):
    """'B8: 1 → 2; C8: ₡340,000.00 → ₡360,000.00', con lo que el robot dice que escribio."""
    partes = []
    for cambio in resultado.get("cambios", []):
        como = _colones if cambio["celda"][0] in dinero else (lambda v: "vacío" if v in (None, "") else _cantidad(v))
        antes = "vacío" if cambio["antes"] in (None, "") else como(cambio["antes"])
        partes.append(f"{cambio['celda']}: {antes} → {como(cambio['despues'])}")
    return "; ".join(partes)


def _guardar(descripcion, funcion, destino, titulo, dinero=(), nota=""):
    """Una escritura de configuracion: llama al robot, mide el tiempo y explica el resultado."""
    try:
        resultado, segundos = datos.escribir(f"SaaS {descripcion} ({session['rol']})", funcion)
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(url_for(destino))
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"{titulo}: los datos ya eran esos. No se escribió nada.", "aviso")
    else:
        flash(f"{titulo}. {_cambios(resultado, dinero)} (hoja {resultado['hoja'].strip()}). {nota}"
              f"Guardado en el Excel en {segundos} s.", "ok")
    for aviso in resultado.get("avisos", []):
        flash(aviso, "aviso")
    return redirect(url_for(destino))


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
    for proporcion in producto["proporciones"]:
        proporcion.update(texto=_por_ciento(proporcion["valor"]), texto_otro=_por_ciento(proporcion["valor_otro"]),
                          texto_maximo=_por_ciento(proporcion["maximo"]))
    otras = [(texto, vista) for otro, (texto, vista) in PANTALLAS_COSTO.items() if otro != codigo]
    return render_template("costo.html", libro=libro, titulo=titulo, p=producto, hornada=hornada, cifras=cifras,
                           otras=otras)


@grupo4.route("/costo-pan-baguette")
@solo_administracion
def costo_baguette():
    return _pantalla_costo("P01")


@grupo4.route("/costo-cupcakes")
@solo_administracion
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


@grupo4.route("/costo/proporcion", methods=["POST"])
@solo_administracion
def proporcion():
    codigo = str(request.form.get("producto") or "").strip()[:10]
    concepto, porcentaje = _dato("concepto", 30), _dato("porcentaje", 12)
    destino = PANTALLAS_COSTO.get(codigo, PANTALLAS_COSTO["P01"])[1]
    if not codigo or not concepto or not porcentaje or porcentaje.startswith("="):
        flash("Elige qué proporción cambiar y escribe el porcentaje. No se escribió nada.", "error")
        return redirect(url_for(destino))
    try:
        resultado, segundos = datos.escribir(
            f"SaaS proporcion {concepto} {codigo} ({session['rol']})",
            lambda: cs.actualizar_proporcion(codigo, concepto, porcentaje, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001
        _avisar(error)
        return redirect(url_for(destino))
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"{resultado['etiqueta']} en {resultado['producto']} ya era {resultado['porcentaje']:g} %. "
              "No se escribió nada.", "aviso")
    else:
        flash(f"{resultado['etiqueta']} en {resultado['producto']}: {resultado['porcentaje_antes']:g} % → "
              f"{resultado['porcentaje']:g} % (hoja {resultado['hoja']}, celda {resultado['celda']}). Costo unitario: "
              f"{_colones(resultado['costo_unitario_antes'])} → {_colones(resultado['costo_unitario'])}. "
              f"Guardado en el Excel en {segundos} s.", "ok")
    for aviso in resultado.get("avisos", []):
        flash(aviso, "aviso")
    return redirect(url_for(destino))


# ============================================================================
# planilla: lectura, cambiar una plaza y agregar una plaza
# ============================================================================
@grupo4.route("/planilla")
@solo_administracion
def planilla():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        leido = cs.leer_planilla(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer la hoja planilla: {error}"), 503
    for fila in leido["plazas"] + leido["resumen"]:
        fila.update(personas_texto=_cantidad(fila["personas"]),
                    **{f"{dato}_texto": _colones(fila.get(dato)) for dato in
                       ("sueldo", "sueldo_total", "cargas", "aguinaldo", "costo_total", "riesgo")})
    return render_template("planilla.html", libro=libro, leido=leido,
                           primera_libre=cs.FILAS_PLAZA_NUEVA[0], ultima_libre=cs.FILAS_PLAZA_NUEVA[-1])


@grupo4.route("/planilla/plaza", methods=["POST"])
@solo_administracion
def cambiar_plaza():
    plaza, personas, sueldo = _dato("plaza", 80), _dato("personas", 6), _dato("sueldo", 15)
    if not plaza or plaza.startswith("="):
        flash("Elige la plaza. No se escribió nada.", "error")
        return redirect(url_for("grupo4.planilla"))
    return _guardar(f"planilla {plaza}", lambda: cs.actualizar_plaza(plaza, personas, sueldo, origen=f"SaaS ({session['rol']})"),
                    "grupo4.planilla", f"Plaza {plaza}", dinero=("C",),
                    nota="El Excel recalculó las cargas sociales, el gasto operacional y, si es panadero o ayudante, el costo unitario. ")


@grupo4.route("/planilla/nueva", methods=["POST"])
@solo_administracion
def nueva_plaza():
    nombre, personas, sueldo = _dato("nombre", 60), _dato("personas", 6), _dato("sueldo", 15)
    if not nombre or personas is None or sueldo is None:
        flash("Escribe el nombre de la plaza, el número de personas y el sueldo. No se escribió nada.", "error")
        return redirect(url_for("grupo4.planilla"))
    return _guardar(f"plaza nueva {nombre}", lambda: cs.agregar_plaza(nombre, personas, sueldo, origen=f"SaaS ({session['rol']})"),
                    "grupo4.planilla", f"Plaza nueva {nombre}", dinero=("C",),
                    nota="Es una plaza administrativa: suma al gasto operacional, no al costo de producción. ")


# ============================================================================
# gasto operacional: lectura y monto de un rubro manual en un mes
# ============================================================================
@grupo4.route("/gasto-operacional")
@solo_administracion
def gasto_operacional():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        leido = cs.leer_gastos(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer la hoja gasto operacional: {error}"), 503
    for rubro in leido["rubros"] + [leido["total"]]:
        rubro.update(meses_texto=[_colones(valor) for valor in rubro["meses"]], anual_texto=_colones(rubro["anual"]))
    meses = [(numero, corto, largo.capitalize()) for numero, (corto, largo) in enumerate(zip(MESES_PANTALLA, leido["meses"]), 1)]
    return render_template("gasto_operacional.html", libro=libro, leido=leido, meses=meses,
                           manuales=[r for r in leido["rubros"] if r["manual"]])


@grupo4.route("/gasto-operacional/guardar", methods=["POST"])
@solo_administracion
def guardar_gasto():
    rubro, mes, monto = _dato("rubro", 60), _dato("mes", 12), _dato("monto", 15)
    if not rubro or not mes or monto is None or rubro.startswith("="):
        flash("Elige el rubro y el mes y escribe el monto. No se escribió nada.", "error")
        return redirect(url_for("grupo4.gasto_operacional"))
    nombre_mes = cs.MESES[int(mes) - 1] if mes.isdigit() and 1 <= int(mes) <= 12 else mes
    return _guardar(f"gasto {rubro} mes {mes}", lambda: cs.actualizar_gasto(rubro, mes, monto, origen=f"SaaS ({session['rol']})"),
                    "grupo4.gasto_operacional", f"{rubro} de {nombre_mes}", dinero=tuple(cs.COLUMNAS_MES),
                    nota="El Excel recalculó el total del mes y el del año. ")


# ============================================================================
# valor activos: lectura, agregar un activo y cambiarlo (dar de baja = cantidad 0)
# ============================================================================
@grupo4.route("/valor-activos")
@con_sesion
def valor_activos():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        leido = cs.leer_activos(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer la hoja valor activos: {error}"), 503
    for activo in leido["activos"]:
        activo.update(cantidad_texto=_cantidad(activo["cantidad"]), vida_texto=_cantidad(activo["vida_util"]),
                      **{f"{dato}_texto": _colones(activo[dato]) for dato in
                         ("precio", "costo", "depreciacion_anual", "depreciacion_mensual", "valor_12_meses")})
    totales = {dato: _colones(valor) for dato, valor in leido["totales"].items() if not dato.startswith("celda_")}
    return render_template("valor_activos.html", libro=libro, leido=leido, totales=totales)


@grupo4.route("/valor-activos/activo", methods=["POST"])
@solo_administracion
def cambiar_activo():
    activo = _dato("activo", 80)
    cantidad, precio, vida = _dato("cantidad", 6), _dato("precio", 15), _dato("vida_util", 6)
    if not activo or activo.startswith("="):
        flash("Elige el activo. No se escribió nada.", "error")
        return redirect(url_for("grupo4.valor_activos"))
    return _guardar(f"activo {activo}",
                    lambda: cs.actualizar_activo(activo, cantidad, precio, vida, origen=f"SaaS ({session['rol']})"),
                    "grupo4.valor_activos", f"Activo {activo}", dinero=("C",),
                    nota="El Excel recalculó el costo, la depreciación y el total de activos. ")


@grupo4.route("/valor-activos/nuevo", methods=["POST"])
@solo_administracion
def nuevo_activo():
    nombre = _dato("nombre", 60)
    cantidad, precio, vida = _dato("cantidad", 6), _dato("precio", 15), _dato("vida_util", 6)
    if not nombre or cantidad is None or precio is None or vida is None:
        flash("Escribe el nombre, la cantidad, el precio y la vida útil. No se escribió nada.", "error")
        return redirect(url_for("grupo4.valor_activos"))
    return _guardar(f"activo nuevo {nombre}",
                    lambda: cs.agregar_activo(nombre, cantidad, precio, vida, origen=f"SaaS ({session['rol']})"),
                    "grupo4.valor_activos", f"Activo nuevo {nombre}", dinero=("C",),
                    nota="El Excel recalculó el total de activos, que es la inversión que lee el flujo de caja del proyecto. ")
