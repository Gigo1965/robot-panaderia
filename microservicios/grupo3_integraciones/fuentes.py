"""Fuentes externas de NODO INTEGRAL (solo LEEN de internet; no tocan el Excel).

Cada funcion 'obtener_...' consulta internet y cada funcion 'interpretar_...'
convierte la respuesta en datos simples. Estan separadas para poder probar la
interpretacion con respuestas de ejemplo, sin internet.

Fuentes (libro de mapas, hoja NODO INTEGRAL):
  - Clima y temperatura de Cañas: API Open-Meteo (sin clave)
  - Tipo de cambio: REFERENCIA OFICIAL DEL BCCR, leida de la API publica de Hacienda.
    Decision del usuario (23-sep-2026): la tabla de ventanilla del BCCR (Banco Nacional)
    fue retirada (404) y el sitio nuevo del BCCR carga los datos con JavaScript. Se
    verifico que Hacienda devuelve EXACTAMENTE el mismo valor que la pagina oficial del
    BCCR (23-sep-2026: compra 447,19 / venta 451,86). Es solo informativo (no mueve costos).
  - Feriados de Costa Rica: libreria Python 'holidays'
  - Dias de pago: regla de negocio (ver es_dia_de_pago)
"""
import calendar
import json
import urllib.request
from datetime import date

URL_CLIMA = ("https://api.open-meteo.com/v1/forecast?latitude=10.4318&longitude=-85.0983"
             "&current=temperature_2m,relative_humidity_2m,apparent_temperature,precipitation,rain,"
             "weather_code,wind_speed_10m&timezone=America%2FCosta_Rica")
URL_HACIENDA = "https://api.hacienda.go.cr/indicadores/tc/dolar"
NAVEGADOR = {"User-Agent": "Mozilla/5.0 (robot-panaderia-canas; proyecto academico)"}

# Codigos de tiempo de Open-Meteo (WMO) que significan lluvia, llovizna o tormenta
CODIGOS_LLUVIA = set(range(51, 68)) | set(range(80, 83)) | set(range(95, 100))


class FuenteNoDisponible(Exception):
    """La fuente no respondio o respondio algo que no se pudo interpretar."""


def _descargar(url, segundos=20):
    solicitud = urllib.request.Request(url, headers=NAVEGADOR)
    try:
        with urllib.request.urlopen(solicitud, timeout=segundos) as respuesta:
            return respuesta.read().decode("utf-8", errors="replace")
    except Exception as error:
        raise FuenteNoDisponible(f"No se pudo leer {url}: {error}") from error


# ---------------------------------------------------------------------------- clima
def interpretar_clima(datos):
    """'Lluvia' si hay precipitacion > 0 o un codigo de lluvia/tormenta; si no, 'Despejado'."""
    try:
        actual = datos["current"]
        temperatura = round(float(actual["temperature_2m"]), 1)
    except (KeyError, TypeError, ValueError) as error:
        raise FuenteNoDisponible(f"Respuesta de clima inesperada: {error}") from error
    precipitacion = float(actual.get("precipitation") or 0) + float(actual.get("rain") or 0)
    codigo = int(actual.get("weather_code") or 0)
    estado = "Lluvia" if precipitacion > 0 or codigo in CODIGOS_LLUVIA else "Despejado"
    return {"estado": estado, "temperatura": temperatura, "codigo": codigo,
            "precipitacion": precipitacion, "hora": actual.get("time")}


def obtener_clima():
    try:
        datos = json.loads(_descargar(URL_CLIMA))
    except json.JSONDecodeError as error:
        raise FuenteNoDisponible(f"El clima no devolvio JSON: {error}") from error
    return interpretar_clima(datos)


# ---------------------------------------------------------------------------- tipo de cambio
def interpretar_tipo_cambio_hacienda(datos):
    """{'venta': {'fecha': ..., 'valor': 451.86}, 'compra': {...}} -> compra, venta y fecha."""
    try:
        compra, venta = float(datos["compra"]["valor"]), float(datos["venta"]["valor"])
    except (KeyError, TypeError, ValueError) as error:
        raise FuenteNoDisponible(f"Respuesta de Hacienda inesperada: {error}") from error
    if not (300 <= venta <= 1500 and 300 <= compra <= venta):
        raise FuenteNoDisponible(f"Tipo de cambio fuera de rango: compra {compra}, venta {venta}")
    fecha = str((datos.get("venta") or {}).get("fecha") or "")[:10]
    return {"compra": compra, "venta": venta, "fecha": fecha,
            "fuente": "Referencia BCCR (API de Hacienda)"}


def obtener_tipo_cambio():
    """Tipo de cambio de referencia del BCCR, del dia de hoy (la API de Hacienda da el vigente)."""
    try:
        return interpretar_tipo_cambio_hacienda(json.loads(_descargar(URL_HACIENDA)))
    except json.JSONDecodeError as error:
        raise FuenteNoDisponible(f"Hacienda no devolvio JSON: {error}") from error


# ---------------------------------------------------------------------------- calendario
def nombre_feriado(fecha: date):
    """Nombre del feriado de Costa Rica en esa fecha, o None si no es feriado.

    Decision del usuario (23-sep-2026): el 5% de feriado aplica SOLO a los 9 feriados de
    PAGO OBLIGATORIO, que es justo lo que trae la libreria 'holidays' (verificado contra el
    calendario 2026 del MTSS). Los 3 de pago no obligatorio (2 ago, 31 ago, 1 dic) NO cuentan.
    """
    try:
        import holidays
    except ImportError as error:
        raise FuenteNoDisponible("Falta la libreria 'holidays': py -m pip install -r requirements.txt") from error
    try:
        feriados = holidays.country_holidays("CR", years=fecha.year, language="es")
    except Exception:
        feriados = holidays.country_holidays("CR", years=fecha.year)
    return feriados.get(fecha)


def es_dia_de_pago(fecha: date):
    """Regla de negocio (confirmada por el usuario, 23-sep-2026):

    - Se paga el dia 15 y el dia 30 de cada mes.
    - Febrero no tiene dia 30: su segundo pago es el ULTIMO dia del mes (28 o 29).
    - Si el dia cae sabado o domingo, igual cuenta ese mismo dia.
    """
    if fecha.day == 15:
        return True
    if fecha.month == 2:
        ultimo_dia = calendar.monthrange(fecha.year, 2)[1]
        return fecha.day == ultimo_dia
    return fecha.day == 30
