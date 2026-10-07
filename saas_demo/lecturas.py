"""SaaS de demostracion - dos hojas que solo se LEEN (6-oct-2026), segun el libro de mapas de celdas.

Regla #39 y decision del usuario (6-oct-2026): todas las hojas del Excel separadas en microservicios
deben VERSE en el SaaS. Estas dos no tenian pantalla propia porque nadie las digita:

  Detalle comanda digital (Grupo 1) -> el mapa: "generado por el robot al recibir la comanda del bot de
      WhatsApp". Una fila por producto de cada pedido; la anulacion agrega filas de rastro 'WA-xxx-A' con
      ESTADO 'ANULADO' (Regla #30: ya no se anula un producto suelto). Hasta hoy solo se veia dentro de
      cada pedido. La ven despacho y administracion (decision del usuario).
  Consumo de Materia Prima (Grupo 2) -> el mapa: la receta (TablaRecetaConsumo) "no se digita: sigue a
      las hojas de costo"; el consumo por hornada (TablaConsumoMP) "lo escribe el robot (hornadas.py),
      valores fijos". La ve solo administracion (decision del usuario).

Aqui NO hay formularios ni funciones de escritura: las dos pantallas leen el Excel tal como esta y lo
muestran. Las tablas se ubican por su nombre y sus columnas por el encabezado (las mismas constantes
que usan pedidos_digitales.py y hornadas.py).
"""
import logging
from datetime import datetime

from flask import Blueprint, render_template, request

from . import config                      # noqa: F401  PRIMERO (prepara el entorno del robot)
from . import datos                       # noqa: E402
from .seguridad import con_sesion, solo_administracion                     # noqa: E402
from conexion_drive import ErrorConexionExcel                              # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla                # noqa: E402
from microservicios.grupo1_captura_operativa import pedidos_digitales as pd   # noqa: E402
from microservicios.grupo2_inventario import hornadas as hr                # noqa: E402

log = logging.getLogger("saas_demo")
lecturas = Blueprint("lecturas", __name__)

LINEAS_EN_PANTALLA = 120                  # lineas de comanda mas recientes que se muestran
CONSUMOS_EN_PANTALLA = 60                 # filas de consumo mas recientes (10 hornadas de 6 materias primas)


def _limpio(valor):
    return " ".join(str(valor if valor is not None else "").split())


def _numero(valor):
    return valor if isinstance(valor, (int, float)) and not isinstance(valor, bool) else None


def _libro_o_problema():
    try:
        return datos.leer_libro(forzar=request.args.get("actualizar") == "1"), None
    except datos.CopiaNoVerificada as error:
        log.error("Copia no verificada: %s", error)
        return None, str(error)
    except (ErrorConexionExcel, KeyError) as error:
        log.exception("No se pudo leer el Excel")
        return None, f"No se pudo leer el Excel de demostración: {error}"


# ============================================================================
# Detalle comanda digital: una fila por producto de cada pedido (solo lectura)
# ============================================================================
def leer_detalle_comanda(wb):
    """Las lineas de 'Detalle comanda digital', la mas reciente primero (wb = libro de VALORES)."""
    ws, tabla = obtener_tabla(wb, pd.HOJA_DETALLE, pd.TABLA_DETALLE)
    lineas = []
    for fila in leer_tabla(ws, tabla):
        id_pedido = _limpio(fila.get("ID_PEDIDO"))
        if not id_pedido:
            continue
        rastro = id_pedido.endswith(pd.SUFIJO_ANULACION)
        descuento = _numero(fila.get("DESCUENTO_ASIGNADO"))
        lineas.append({
            "fila": fila["_fila"], "id": id_pedido,
            "pedido": id_pedido[:-len(pd.SUFIJO_ANULACION)] if rastro else id_pedido,      # para el enlace al pedido
            "rastro": rastro, "fecha": fila.get("FECHA") if isinstance(fila.get("FECHA"), datetime) else None,
            "producto": _limpio(fila.get("PRODUCTO")), "cantidad": _numero(fila.get("CANTIDAD")),
            "estado": _limpio(fila.get("ESTADO")), "subtotal": _numero(fila.get("SUBTOTAL_LINEA")),
            "oferta": _limpio(fila.get("CODIGO_OFERTA")), "detonador": _limpio(fila.get("TIPO_DETONADOR")),
            "descuento": f"{round(descuento * 100, 2):g} %" if descuento else "",
        })
    lineas.sort(key=lambda linea: linea["fila"], reverse=True)       # lo ultimo que se escribio, arriba
    return lineas


@lecturas.route("/detalle-comanda-digital")
@con_sesion
def detalle_comanda():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        lineas = leer_detalle_comanda(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer Detalle comanda digital: {error}"), 503
    cuentas = {"confirmadas": sum(1 for l in lineas if l["estado"] == pd.CONFIRMADO),
               "anuladas": sum(1 for l in lineas if l["estado"] == pd.ANULADO),
               "pedidos": len({l["pedido"] for l in lineas})}
    return render_template("detalle_comanda.html", libro=libro, lineas=lineas[:LINEAS_EN_PANTALLA], total=len(lineas),
                           cuentas=cuentas, hoja=pd.HOJA_DETALLE)


# ============================================================================
# Consumo de Materia Prima: la receta y lo que consumio cada hornada (solo lectura)
# ============================================================================
def leer_consumo_mp(wb):
    """{'receta': [...], 'consumos': [...mas reciente primero...], 'hornadas': n} (wb = libro de VALORES)."""
    ws_r, tabla_r = obtener_tabla(wb, hr.HOJA_CONSUMO, hr.TABLA_RECETA)
    receta = [{"fila": f["_fila"], "producto": _limpio(f.get("Producto horneado")), "codigo": _limpio(f.get("Codigo")),
               "materia": _limpio(f.get("Materia prima")), "cantidad": _numero(f.get("Cantidad por unidad")),
               "unidad": _limpio(f.get("Unidad")), "contenido": _numero(f.get("Contenido por paquete"))}
              for f in leer_tabla(ws_r, tabla_r) if _limpio(f.get("Materia prima"))]
    ws_c, tabla_c = obtener_tabla(wb, hr.HOJA_CONSUMO, hr.TABLA_CONSUMO)
    consumos = []
    for f in leer_tabla(ws_c, tabla_c):
        if not _limpio(f.get("Lote")):
            continue
        horneado = _numero(f.get("Cantidad horneada"))
        consumos.append({"fila": f["_fila"], "fecha": f.get("Fecha") if isinstance(f.get("Fecha"), datetime) else None,
                         "lote": _limpio(f.get("Lote")), "producto": _limpio(f.get("Producto horneado")),
                         "horneado": horneado, "materia": _limpio(f.get("Materia prima")),
                         "por_unidad": _numero(f.get("Consumo por unidad")), "total": _numero(f.get("Consumo total")),
                         "contenido": _numero(f.get("Contenido por paquete")),
                         "paquetes": _numero(f.get("Paquetes consumidos")),
                         "correccion": horneado is not None and horneado < 0})
    consumos.sort(key=lambda fila: fila["fila"], reverse=True)       # lo ultimo que se escribio, arriba
    return {"receta": receta, "consumos": consumos, "hornadas": len({c["lote"] for c in consumos})}


@lecturas.route("/consumo-materia-prima")
@solo_administracion
def consumo_mp():
    libro, problema = _libro_o_problema()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    try:
        leido = leer_consumo_mp(libro["wb"])
    except KeyError as error:
        return render_template("problema.html", problema=f"No se pudo leer Consumo de Materia Prima: {error}"), 503
    return render_template("consumo_materia_prima.html", libro=libro, receta=leido["receta"],
                           consumos=leido["consumos"][:CONSUMOS_EN_PANTALLA], total=len(leido["consumos"]),
                           hornadas=leido["hornadas"], hoja=hr.HOJA_CONSUMO)
