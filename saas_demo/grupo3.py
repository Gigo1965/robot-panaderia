"""SaaS de demostracion - Grupo 3 (Integraciones), segun el libro de mapas de celdas.

Un modulo por hoja, en el orden del Excel maestro:
  NODO INTEGRAL           -> nodo_manual.py (celdas manuales C3, C4, C8 y C22, Decisiones Pendientes #64).
                             activacion_promocion.py (C16 'Activacion de Promocion', #84): la escribe el
                             ROBOT; la pantalla solo muestra "Oferta hoy: ON / OFF" (lectura).
                             || 4-oct-2026 (noche): el usuario quito el boton "Actualizar indicador de oferta":
                             C16 es una celda bloqueada segun el mapa y nadie la maneja a mano. La escribe el
                             robot (tarea diaria en el maestro; al guardar un combo en la copia, paquete 3b).
                             Las demas filas las escribe el robot por horario o son formulas (Regla #4):
                             se muestran como LECTURA. C14 y C15 quedan solo lectura (Regla #8, #80).
  TABLA CRUCE MAESTRA     -> el mapa no marca ninguna celda manual: LECTURA.
  TABLA CRUCE MAESTRA MP  -> el mapa no marca ninguna celda manual: LECTURA.
  Tabla Combos Dinamicos  -> paquete 3b (Decisiones Pendientes #65), todavia sin construir.
                             || 5-oct-2026: combos_config.py -> la CONFIGURACION de cada combo (G tipo de
                             detonador, H operador, I valor limite, J prioridad, K mensaje y L descuento),
                             un formulario por combo, solo administracion. La columna B la reescribe el
                             robot y, despues de guardar, el robot revisa C16. La COMPOSICION del combo
                             (que productos lleva, #85) todavia no se cambia desde aqui.
                             || 5-oct-2026 (tarde): combos_composicion.py -> la COMPOSICION de cada combo (#85):
                             un producto de produccion propia y uno de reventa, con su cantidad (1 a 10). Segundo
                             formulario por combo, solo administracion. El robot congela lo ya rebajado del
                             inventario (migracion m014, Regla #40) y reescribe la formula del precio, el
                             detalle y, si trae el detalle viejo, el mensaje.

Aqui NO hay reglas de negocio: la pantalla recoge lo digitado y se lo entrega a la funcion del
robot (actualizar_dato_manual). Si el dato no se permite, lo rechaza el robot y la pantalla
muestra su mensaje.

Permisos (decision del usuario, 4-oct-2026): las lecturas las ven los dos usuarios; las celdas
manuales las cambia solo administracion.
"""
import logging

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from . import config                      # noqa: F401  PRIMERO (prepara el entorno del robot)
from . import datos                       # noqa: E402
from .seguridad import con_sesion, solo_administracion                     # noqa: E402
from conexion_drive import ConflictoDeVersion, ErrorConexionExcel          # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla                # noqa: E402
from microservicios.grupo3_integraciones import activacion_promocion as ap  # noqa: E402
from microservicios.grupo3_integraciones import combos_composicion as co   # noqa: E402
from microservicios.grupo3_integraciones import combos_config as cc        # noqa: E402
from microservicios.grupo3_integraciones import nodo_manual as nm          # noqa: E402

log = logging.getLogger("saas_demo")
grupo3 = Blueprint("grupo3", __name__)

BOTON = {"tarifa_km": "Guardar tarifa", "meta_venta": "Guardar meta", "meta_ads": "Guardar acción",
         "sinpe": "Guardar número"}

# Las dos tablas cruce maestras: (encabezado en el Excel, titulo en pantalla, como se muestra).
TABLAS_CRUCE = {
    "tcm": {
        "hoja": "TABLA CRUCE MAESTRA", "tabla": "TablaProductos", "llave": "Nombre tecnico completo",
        "titulo": "Tabla cruce maestra", "vista": "grupo3.tabla_cruce_maestra",
        "columnas": [("Codigo", "Código", "texto"), ("Nombre tecnico completo", "Nombre técnico (llave)", "texto"),
                     ("Nombre visible", "Nombre visible", "texto"), ("Costo Unitario (Robot)", "Costo unitario", "colones"),
                     ("Precio", "Precio de venta", "colones"), ("Fuente de Costo", "Fuente del costo", "texto"),
                     ("Tipo", "Tipo", "texto")],
    },
    "mp": {
        "hoja": "TABLA CRUCE MAESTRA MP", "tabla": "TablaCruceMaestraMP", "llave": "Materia Prima",
        "titulo": "Tabla cruce maestra MP", "vista": "grupo3.tabla_cruce_mp",
        "columnas": [("Codigo", "Código", "texto"), ("Materia Prima", "Materia prima (llave)", "texto"),
                     ("Nombre Visible", "Nombre visible", "texto"),
                     ("Costo Unitario (Scraping)", "Costo del paquete", "colones"),
                     ("Por gramo o pieza", "Costo por gramo o pieza", "colones4"), ("Fuente", "Fuente", "texto"),
                     ("Tipo", "Tipo", "texto")],
    },
}


def _numero(valor):
    return valor if isinstance(valor, (int, float)) and not isinstance(valor, bool) else None


def _colones(valor, decimales=2):
    """500 -> '₡500', 175000.5 -> '₡175,000.50', 0.7004 (4 decimales) -> '₡0.7004'."""
    if _numero(valor) is None:
        return "" if valor in (None, "") else str(valor)
    if decimales == 4:
        return f"₡{valor:,.4f}"
    return f"₡{valor:,.0f}" if round(valor, 2) == int(valor) else f"₡{valor:,.2f}"


def _texto_del_valor(fila):
    """Como se muestra en pantalla el valor de una fila de NODO INTEGRAL."""
    valor = fila["valor"]
    if fila["campo"] and nm.CAMPOS[fila["campo"]]["tipo"] == "colones":
        return _colones(valor)
    if _numero(valor) is not None:
        return f"{valor:,.2f}".rstrip("0").rstrip(".")
    return "" if valor is None else str(valor).strip()


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
    if isinstance(error, (nm.ValorInvalido, cc.ValorInvalido, ap.ActivacionInvalida, LookupError, ValueError)):
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
# NODO INTEGRAL: lectura de toda la hoja y celdas manuales (#64)
# ============================================================================
@grupo3.route("/nodo-integral")
@con_sesion
def nodo_integral():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        filas = nm.leer_nodo(libro["wb"])
        oferta = ap.leer_activacion(libro["wb"])
    except (KeyError, ap.ActivacionInvalida) as error:
        return render_template("problema.html", problema=f"No se pudo leer NODO INTEGRAL: {error}"), 503
    for fila in filas:
        fila["texto"] = _texto_del_valor(fila)
    manuales = [dict(f, definicion=nm.CAMPOS[f["campo"]], boton=BOTON[f["campo"]]) for f in filas if f["clase"] == "manual"]
    return render_template("nodo_integral.html", libro=libro, filas=filas, manuales=manuales, oferta=oferta)


@grupo3.route("/nodo-integral/guardar", methods=["POST"])
@solo_administracion
def guardar_dato():
    campo = str(request.form.get("campo") or "").strip()[:30]
    valor = str(request.form.get("valor") or "").strip()[:40]
    if not campo or not valor or valor.startswith("="):
        flash("Escribe el valor nuevo antes de guardar. No se escribió nada.", "error")
        return redirect(url_for("grupo3.nodo_integral"))
    try:
        resultado, segundos = datos.escribir(
            f"SaaS NODO INTEGRAL {campo} ({session['rol']})",
            lambda: nm.actualizar_dato_manual(campo, valor, origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(url_for("grupo3.nodo_integral"))
    es_monto = nm.CAMPOS[resultado["campo"]]["tipo"] == "colones"
    mostrar = _colones if es_monto else (lambda v: "" if v is None else str(v).strip())
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"{resultado['etiqueta']} ya era {mostrar(resultado['valor'])}. No se escribió nada.", "aviso")
    else:
        flash(f"{resultado['etiqueta']}: {mostrar(resultado['antes'])} → {mostrar(resultado['despues'])} "
              f"(celda {resultado['celda']}). Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo3.nodo_integral"))


# ============================================================================
# TABLA CRUCE MAESTRA y TABLA CRUCE MAESTRA MP: solo lectura
# ============================================================================
def _tabla_cruce(clave):
    definicion = TABLAS_CRUCE[clave]
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        ws, tabla = obtener_tabla(libro["wb"], definicion["hoja"], definicion["tabla"])
        registros = [r for r in leer_tabla(ws, tabla) if str(r.get(definicion["llave"]) or "").strip()]
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer {definicion['hoja']}: {error}"), 503
    filas = []
    for registro in registros:
        celdas = []
        for encabezado, titulo, forma in definicion["columnas"]:
            valor = registro.get(encabezado)
            if forma == "texto":
                celdas.append({"titulo": titulo, "texto": "" if valor is None else str(valor).strip(), "num": False})
            else:
                celdas.append({"titulo": titulo, "texto": _colones(valor, 4 if forma == "colones4" else 2), "num": True})
        filas.append({"fila": registro["_fila"], "celdas": celdas})
    return render_template("tabla_cruce.html", libro=libro, clave=clave, definicion=definicion, filas=filas,
                           rango=tabla.ref)


@grupo3.route("/tabla-cruce-maestra")
@con_sesion
def tabla_cruce_maestra():
    return _tabla_cruce("tcm")


@grupo3.route("/tabla-cruce-maestra-mp")
@con_sesion
def tabla_cruce_mp():
    return _tabla_cruce("mp")


# ============================================================================
# Tabla Combos Dinamicos: configuracion de cada combo (#65)
# ============================================================================
def _criterio(combo):
    """'Temperatura <= 30 °C' o 'Clima = Lluvia', como se lee en pantalla."""
    if combo["tipo"] == cc.TEMPERATURA:
        valor = f"{combo['valor']:g}" if _numero(combo["valor"]) is not None else str(combo["valor"])
        return f"Temperatura {combo['operador']} {valor} °C"
    if combo["tipo"] == cc.LLUVIA:
        return f"Clima = {combo['valor']}"
    return f"{combo['tipo']} {combo['operador']} {combo['valor']}".strip()


def _valor_en_pantalla(valor):
    if _numero(valor) is not None:
        return f"{valor:g}"
    return "" if valor is None else str(valor).strip()


@grupo3.route("/combos")
@con_sesion
def combos():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        lista = cc.leer_combos(libro["wb"])
        oferta = ap.leer_activacion(libro["wb"])
        nodo = {f["descripcion"]: f["valor"] for f in nm.leer_nodo(libro["wb"])}
    except (KeyError, cc.ValorInvalido, ap.ActivacionInvalida) as error:
        return render_template("problema.html", problema=f"No se pudo leer Tabla Combos Dinamicos: {error}"), 503
    for combo in lista:
        combo["criterio"] = _criterio(combo)
        combo["precio_texto"] = _colones(combo["precio"])
        combo["temperatura"] = f"{combo['valor']:g}" if _numero(combo["valor"]) is not None else ""
        combo["lluvia"] = combo["valor"] if combo["valor"] in cc.CLIMAS else ""
    hoy = {"clima": _valor_en_pantalla(nodo.get("Clima Actual")),
           "temperatura": _valor_en_pantalla(nodo.get("Temperatura clima")),
           "musi": _valor_en_pantalla(nodo.get("DATA OFERTA MUSI"))}
    # La composicion (#85) se lee aparte: si esa tabla no se puede leer, el resto de la pantalla sigue funcionando.
    try:
        leido = co.leer_composicion(libro["wb"])
        composicion = {"combos": {c["codigo"]: c for c in leido["combos"]}, "propios": leido["propios"],
                       "terceros": leido["terceros"], "problema": ""}
    except (KeyError, cc.ValorInvalido) as error:
        composicion = {"combos": {}, "propios": [], "terceros": [], "problema": str(error)}
    return render_template("combos.html", libro=libro, combos=lista, oferta=oferta, hoy=hoy, cc=cc,
                           composicion=composicion, co=co)


@grupo3.route("/combos/guardar", methods=["POST"])
@solo_administracion
def guardar_combo():
    def dato(nombre, largo):
        return str(request.form.get(nombre) or "").strip()[:largo]

    codigo, tipo = dato("codigo", 12), dato("tipo", 20)
    # La pantalla trae los dos campos del valor limite; se entrega el que corresponde al detonador elegido.
    valor = dato("lluvia", 20) if tipo.upper() == cc.LLUVIA else dato("temperatura", 12)
    try:
        resultado, segundos = datos.escribir(
            f"SaaS combo {codigo} ({session['rol']})",
            lambda: cc.actualizar_combo(codigo, tipo, dato("operador", 4), valor, dato("prioridad", 6),
                                        dato("mensaje", 400), dato("descuento", 12),
                                        origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(url_for("grupo3.combos"))
    titulo = f"{resultado['nombre']} ({resultado['codigo']})"
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"{titulo}: la configuración ya era esa. No se escribió nada.", "aviso")
        return redirect(url_for("grupo3.combos"))
    cambios = "; ".join(f"{c['titulo']} {_valor_en_pantalla(c['antes'])} → {_valor_en_pantalla(c['despues'])} (celda {c['celda']})"
                        for c in resultado["cambios"])
    activacion = resultado.get("activacion") or {}
    oferta = f" Oferta hoy pasó a {activacion['valor']}." if activacion.get("hubo_cambios") else ""
    flash(f"{titulo}: {cambios}.{oferta} Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo3.combos"))


# ============================================================================
# Tabla Combos Dinamicos: composicion de cada combo (#85)
# ============================================================================
def _cantidad_en_pantalla(valor):
    return f"{valor:g}" if _numero(valor) is not None else str(valor)


@grupo3.route("/combos/composicion/guardar", methods=["POST"])
@solo_administracion
def guardar_composicion():
    def dato(nombre, largo):
        return str(request.form.get(nombre) or "").strip()[:largo]

    codigo = dato("combo", 12)
    try:
        resultado, segundos = datos.escribir(
            f"SaaS composicion {codigo} ({session['rol']})",
            lambda: co.actualizar_composicion(codigo, dato("producto_1", 120), dato("cantidad_1", 6),
                                              dato("producto_2", 120), dato("cantidad_2", 6),
                                              origen=f"SaaS ({session['rol']})"))
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar)
        _avisar(error)
        return redirect(url_for("grupo3.combos"))
    titulo = f"{resultado['nombre']} ({resultado['codigo']})"
    if resultado.get("estado") == "SIN_CAMBIO":
        flash(f"{titulo}: la composición ya era esa ({resultado['antes']}). No se escribió nada.", "aviso")
        return redirect(url_for("grupo3.combos"))
    congelado = "; ".join(f"{_cantidad_en_pantalla(d['unidades'])} {d['producto']}" for d in resultado["congelado"]
                          if d["unidades"])
    congelado = f" Lo ya rebajado del inventario queda congelado: {congelado}." if congelado else ""
    mensaje = {"SIN_DETALLE": " El mensaje del combo no traía el detalle anterior y no se cambió: revísalo en su configuración.",
               "MUY_LARGO": " El mensaje del combo no se cambió porque pasaría del largo permitido: revísalo en su configuración.",
               }.get(resultado.get("mensaje"), "")
    flash(f"{titulo}: {resultado['antes']} → {resultado['despues']}.{congelado}{mensaje} "
          f"Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("grupo3.combos"))
