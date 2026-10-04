"""Datos del SaaS de demostracion: leer los pedidos y llamar a las funciones COMPARTIDAS.

No tiene reglas de negocio propias (Regla #4 y #5): entregar y anular los resuelve
pedidos_digitales.py, el mismo archivo que usa el bot de WhatsApp. Aqui solo se:
  1) revisa que el archivo sea la COPIA de demostracion (candado 3 de config.py);
  2) lee 'Interface digital' y 'Detalle comanda digital' para mostrarlos;
  3) llama a marcar_entregado / anular_pedido_digital, UNA escritura a la vez,
     y mide cuanto tarda (dato para el reporte).
"""
import logging
import threading
import time
from datetime import datetime, time as hora_del_dia

from . import config                      # PRIMERO: prepara el entorno (ver config.py)

from conexion_drive import config as config_robot                       # noqa: E402
from conexion_drive import leer_excel, obtener_metadatos                # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla             # noqa: E402
from conexion_drive.tiempo import ahora_local                           # noqa: E402
from microservicios.grupo1_captura_operativa import pedidos_digitales as pd   # noqa: E402

log = logging.getLogger("saas_demo")

PENDIENTE, ENTREGADO, ANULADO, OTRO = "pendiente", "entregado", "anulado", "otro"
SEGUNDOS_ESPERA_CANDADO = 100          # menos que el --timeout 120 de gunicorn

_candado_escritura = threading.Lock()   # una escritura a la vez (LibreOffice es pesado)
_candado_cache = threading.Lock()
_cache = {"momento": 0.0, "datos": None}
_verificacion = {"momento": 0.0, "nombre": None}


class CopiaNoVerificada(Exception):
    """El archivo configurado no es la copia de demostracion: no se lee ni se escribe."""


class Ocupado(Exception):
    """Hay otra escritura en curso y no termino a tiempo."""


# ----------------------------------------------------------------------------
# Candado 3: el archivo debe ser la COPIA de demostracion
# ----------------------------------------------------------------------------
def nombre_es_copia(nombre, palabra=None):
    palabra = (palabra or config.PALABRA_COPIA).strip().lower()
    return bool(palabra) and palabra in str(nombre or "").lower()


def verificar_copia(reciente=False):
    """Revisa que el robot este apuntando a la copia. Devuelve el nombre del archivo.

    reciente=True (antes de ESCRIBIR): vuelve a preguntar a Drive siempre.
    reciente=False (para leer): acepta una revision de hace menos de 5 minutos.
    """
    if config.ENSAYO:
        if config_robot.ORIGEN_EXCEL != "local":
            raise CopiaNoVerificada("El modo ENSAYO debe trabajar con una copia local. No se hizo nada.")
        return "copia temporal de ensayo"
    if config_robot.ORIGEN_EXCEL != "drive" or config_robot.EXCEL_MAESTRO_FILE_ID != config.ID_DEMO:
        raise CopiaNoVerificada(
            "El robot no esta apuntando al archivo de demostracion (SAAS_EXCEL_FILE_ID). No se hizo nada.")
    if not reciente and _verificacion["nombre"] and time.monotonic() - _verificacion["momento"] < 300:
        return _verificacion["nombre"]
    nombre = obtener_metadatos().get("name")
    if not nombre_es_copia(nombre):
        raise CopiaNoVerificada(
            f"El archivo de Drive se llama '{nombre}' y no contiene la palabra '{config.PALABRA_COPIA}': "
            "no parece la copia de demostracion. No se leyo ni se escribio nada.")
    _verificacion.update(momento=time.monotonic(), nombre=nombre)
    return nombre


# ----------------------------------------------------------------------------
# Lectura
# ----------------------------------------------------------------------------
def _limpio(valor):
    return str(valor).strip() if valor not in (None, "") else ""


def _numero(valor):
    return float(valor) if isinstance(valor, (int, float)) and not isinstance(valor, bool) else 0.0


def _es_sinpe(valor):
    """Clasifica con la funcion del robot (no se repite la regla aqui)."""
    try:
        return pd.normalizar_metodo_pago(valor) == pd.SINPE
    except pd.MetodoPagoInvalido:
        return False


def _telefono_oculto(valor):
    """'+50688887777' -> 'termina en 7777'. El numero completo no se muestra en pantalla."""
    digitos = "".join(c for c in str(valor or "") if c.isdigit())
    return f"termina en {digitos[-4:]}" if len(digitos) >= 8 else ""


def _orden(pedido):
    fecha = pedido["fecha"] if isinstance(pedido["fecha"], datetime) else datetime.min
    hora = pedido["hora"] if isinstance(pedido["hora"], hora_del_dia) else hora_del_dia.min
    return datetime.combine(fecha.date(), hora)


def _construir(wb):
    ws_i, tabla_i = obtener_tabla(wb, pd.HOJA_INTERFACE, pd.TABLA_INTERFACE)
    ws_d, tabla_d = obtener_tabla(wb, pd.HOJA_DETALLE, pd.TABLA_DETALLE)
    filas = leer_tabla(ws_i, tabla_i)

    lineas = {}
    for linea in leer_tabla(ws_d, tabla_d):
        lineas.setdefault(_limpio(linea.get("ID_PEDIDO")), []).append({
            "cantidad": linea.get("CANTIDAD"),
            "producto": _limpio(linea.get("PRODUCTO")),
            "subtotal": _numero(linea.get("SUBTOTAL_LINEA")),
            "estado": _limpio(linea.get("ESTADO")),
        })

    anulaciones = {}
    for fila in filas:
        id_fila = _limpio(fila.get("ID_PEDIDO"))
        if id_fila.endswith(pd.SUFIJO_ANULACION):
            anulaciones[id_fila[:-len(pd.SUFIJO_ANULACION)]] = fila

    pedidos = []
    for fila in filas:
        id_pedido = _limpio(fila.get("ID_PEDIDO"))
        if not id_pedido or id_pedido.endswith(pd.SUFIJO_ANULACION):
            continue
        estado_excel = _limpio(fila.get("ESTADO"))
        entregado_en = fila.get(pd.COL_ENTREGADO) if isinstance(fila.get(pd.COL_ENTREGADO), datetime) else None
        anulacion = anulaciones.get(id_pedido)
        pago = _limpio(fila.get("METODO_PAGO"))

        if anulacion is not None:
            situacion = ANULADO
        elif entregado_en is not None:
            situacion = ENTREGADO
        elif estado_excel == pd.CONFIRMADO:
            situacion = PENDIENTE
        else:
            situacion = OTRO

        # Solo para EXPLICAR en pantalla por que no hay boton; quien decide es el robot.
        if situacion == ANULADO:
            por_que_no_anula = "Ya está anulado."
        elif situacion == ENTREGADO:
            por_que_no_anula = "Ya se entregó: el cliente recibió el pedido y pagó."
        elif situacion == OTRO:
            por_que_no_anula = "No está confirmado."
        elif _es_sinpe(pago):
            por_que_no_anula = "Se pagó por SINPE Móvil: el pago es en firme."
        else:
            por_que_no_anula = ""

        texto_anulacion = _limpio(anulacion.get("DETALLE_COMANDA")) if anulacion else ""
        motivo = texto_anulacion
        if motivo.startswith("ANULADO:"):
            motivo = motivo[len("ANULADO:"):].rsplit("(anula", 1)[0].strip()

        pedidos.append({
            "id": id_pedido,
            "fecha": fila.get("FECHA"),
            "hora": fila.get("HORA"),
            "fuente": _limpio(fila.get("FUENTE")),
            "cliente": _limpio(fila.get("CLIENTE")),
            "comanda": _limpio(fila.get("DETALLE_COMANDA")),
            "subtotal": _numero(fila.get("SUBTOTAL_BRUTO")),
            "descuento": _numero(fila.get("DESCUENTO_TOTAL")),
            "envio": _numero(fila.get("ENVIO_MAPS")),
            "total": _numero(fila.get("TOTAL_A_COBRAR")),
            "pago": pago,
            "estado_excel": estado_excel,
            "entregado_en": entregado_en,
            "telefono": _telefono_oculto(fila.get(pd.COL_TELEFONO)),
            "situacion": situacion,
            "puede_entregar": situacion == PENDIENTE,
            "puede_anular": situacion == PENDIENTE and not por_que_no_anula,
            "por_que_no_anula": por_que_no_anula,
            "motivo_anulacion": motivo,
            "anulado_a_las": anulacion.get("HORA") if anulacion else None,
            "lineas": [l for l in lineas.get(id_pedido, []) if l["estado"] == pd.CONFIRMADO],
            "fila": fila["_fila"],
        })
    pedidos.sort(key=_orden, reverse=True)
    return pedidos


def leer_libro(forzar=False):
    """El Excel de demostracion ya leido, para TODOS los modulos del SaaS.

    Devuelve {'wb': libro de VALORES (solo lectura), 'archivo', 'version', 'leido'}.
    Se guarda unos segundos en memoria (SAAS_SEGUNDOS_CACHE) y se olvida despues de cada escritura.
    """
    with _candado_cache:
        vigente = _cache["datos"] is not None and time.monotonic() - _cache["momento"] < config.SEGUNDOS_CACHE
        if vigente and not forzar:
            return _cache["datos"]
        archivo = verificar_copia()
        version = None if config.ENSAYO else obtener_metadatos().get("version")
        datos = {"wb": leer_excel(), "archivo": archivo, "version": version, "leido": ahora_local()}
        _cache.update(momento=time.monotonic(), datos=datos)
        return datos


def leer_pedidos(forzar=False):
    """{'pedidos': [...], 'archivo': nombre, 'version': n, 'leido': datetime, 'cuentas': {...}}"""
    libro = leer_libro(forzar)
    if "pedidos" not in libro:
        pedidos = _construir(libro["wb"])
        cuentas = {s: sum(1 for p in pedidos if p["situacion"] == s) for s in (PENDIENTE, ENTREGADO, ANULADO)}
        cuentas["todos"] = len(pedidos)
        libro.update(pedidos=pedidos, cuentas=cuentas)
    return libro


def buscar_pedido(id_pedido, forzar=False):
    id_pedido = str(id_pedido or "").strip().upper()
    return next((p for p in leer_pedidos(forzar)["pedidos"] if p["id"].upper() == id_pedido), None)


def olvidar_cache():
    with _candado_cache:
        _cache.update(momento=0.0, datos=None)


# ----------------------------------------------------------------------------
# Escritura: las DOS funciones compartidas, una a la vez y con tiempo medido
# ----------------------------------------------------------------------------
def escribir(descripcion, funcion):
    """Ejecuta UNA escritura: revisa la copia, una a la vez, mide el tiempo. Devuelve (resultado, segundos)."""
    if not _candado_escritura.acquire(timeout=SEGUNDOS_ESPERA_CANDADO):
        raise Ocupado("Hay otra escritura en curso en el Excel. Espera un momento y vuelve a intentarlo.")
    inicio = time.monotonic()
    try:
        verificar_copia(reciente=True)
        resultado = funcion()
    finally:
        _candado_escritura.release()
        olvidar_cache()
    segundos = round(time.monotonic() - inicio, 1)
    log.info("%s: listo en %s s", descripcion, segundos)
    return resultado, segundos


def entregar(id_pedido, rol):
    """Boton 'Marcar entregado' -> marcar_entregado (la misma funcion del bot)."""
    return escribir(f"SaaS entregar {id_pedido} ({rol})",
                     lambda: pd.marcar_entregado(id_pedido, origen=f"SaaS ({rol})"))


def anular(id_pedido, motivo, rol):
    """Boton 'Anular pedido' -> anular_pedido_digital (solo administracion; lo exige app.py)."""
    return escribir(f"SaaS anular {id_pedido} ({rol})",
                     lambda: pd.anular_pedido_digital(id_pedido, motivo, origen=f"SaaS ({rol})"))
