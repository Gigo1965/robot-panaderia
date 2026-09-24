"""Grupo 3B -- Google Maps: distancia, tiempo y costo de envio por pedido.

Que hace:
    Recibe la ubicacion que el cliente compartio por WhatsApp (latitud y
    longitud), le pregunta a Google Maps (Routes API) cuantos km y minutos hay
    en carro desde la panaderia, y calcula el costo de envio:

        envio = km x Tarifa Base KM   (NODO INTEGRAL C3), sin IVA
        redondeado HACIA ARRIBA al multiplo de 50 colones

    Si el cliente esta mas lejos que el radio maximo, el pedido queda
    "fuera de zona" (no se calcula envio).

Reglas del proyecto:
    - Este modulo es la UNICA puerta a Google Maps. Lo usa la funcion
      compartida de pedidos (bot de WhatsApp y SaaS), nunca cada canal por
      su cuenta. Es la excepcion a "solo NODO INTEGRAL llama APIs": la
      distancia depende de cada cliente, no se puede calcular por horario.
    - La tarifa NO se escribe aqui: se lee del Excel (NODO INTEGRAL C3).

Configuracion (archivo .env de la raiz):
    GOOGLE_MAPS_API_KEY   clave de Google Maps (obligatoria, nunca en el codigo)
    ORIGEN_LAT, ORIGEN_LNG  ubicacion de la panaderia (opcional; por defecto
                            las coordenadas del proyecto academico)
    RADIO_MAXIMO_KM       radio de entrega (opcional; por defecto 10)
    REDONDEO_ENVIO        multiplo de redondeo del envio (opcional; por defecto 50)

Prueba (desde la raiz ROBOT PANADERO, en la Terminal):
    py -m microservicios.grupo3_integraciones.google_maps 10.4400 -85.0900
    (los dos numeros son la latitud y longitud de un "cliente" de prueba)
"""
import json
import math
import os
import sys
import urllib.error
import urllib.request

from conexion_drive import config  # noqa: F401  (carga el archivo .env)

URL_ROUTES = "https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix"
CAMPOS = "originIndex,destinationIndex,duration,distanceMeters,status,condition"

# Coordenadas del proyecto academico (aprox. 100 m al norte del local, Canas).
_ORIGEN_LAT_DEFECTO = 10.42895
_ORIGEN_LNG_DEFECTO = -85.09425


class ErrorGoogleMaps(Exception):
    """Error al calcular la distancia (clave, permisos, red, sin ruta...)."""


def _numero_env(nombre, defecto):
    valor = (os.getenv(nombre) or "").strip()
    if not valor:
        return defecto
    try:
        return float(valor)
    except ValueError as error:
        raise ErrorGoogleMaps(f"{nombre} en el .env no es un numero: {valor!r}") from error


def configuracion():
    """Lee la configuracion vigente (se lee cada vez, por si cambia el .env)."""
    return {
        "origen_lat": _numero_env("ORIGEN_LAT", _ORIGEN_LAT_DEFECTO),
        "origen_lng": _numero_env("ORIGEN_LNG", _ORIGEN_LNG_DEFECTO),
        "radio_maximo_km": _numero_env("RADIO_MAXIMO_KM", 10.0),
        "redondeo": _numero_env("REDONDEO_ENVIO", 50.0),
    }


def _clave():
    clave = (os.getenv("GOOGLE_MAPS_API_KEY") or "").strip()
    if not clave:
        raise ErrorGoogleMaps("Falta GOOGLE_MAPS_API_KEY en el archivo .env")
    return clave


def _validar_coordenadas(lat, lng):
    try:
        lat, lng = float(lat), float(lng)
    except (TypeError, ValueError) as error:
        raise ErrorGoogleMaps(f"Coordenadas invalidas: {lat!r}, {lng!r}") from error
    if not (-90 <= lat <= 90 and -180 <= lng <= 180):
        raise ErrorGoogleMaps(f"Coordenadas fuera de rango: {lat}, {lng}")
    return lat, lng


def _punto(lat, lng):
    return {"waypoint": {"location": {"latLng": {"latitude": lat, "longitude": lng}}}}


def consultar_ruta(destino_lat, destino_lng, timeout=15):
    """Pregunta a Google Maps la ruta en carro panaderia -> cliente.

    Devuelve {"km": float, "minutos": int}. Lanza ErrorGoogleMaps si falla.
    """
    destino_lat, destino_lng = _validar_coordenadas(destino_lat, destino_lng)
    cfg = configuracion()
    cuerpo = {
        "origins": [_punto(cfg["origen_lat"], cfg["origen_lng"])],
        "destinations": [_punto(destino_lat, destino_lng)],
        "travelMode": "DRIVE",
    }
    solicitud = urllib.request.Request(
        URL_ROUTES,
        data=json.dumps(cuerpo).encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/json",
            "X-Goog-Api-Key": _clave(),
            "X-Goog-FieldMask": CAMPOS,
        },
    )
    try:
        with urllib.request.urlopen(solicitud, timeout=timeout) as respuesta:
            datos = json.loads(respuesta.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detalle = error.read().decode("utf-8", "replace")[:300]
        if error.code == 403:
            raise ErrorGoogleMaps(
                "Google nego el permiso (403). Revisa que la clave sea correcta y que "
                f"'Routes API' este habilitada en Google Cloud. Detalle: {detalle}"
            ) from error
        raise ErrorGoogleMaps(f"Google Maps respondio error {error.code}: {detalle}") from error
    except (urllib.error.URLError, TimeoutError) as error:
        raise ErrorGoogleMaps(f"No se pudo conectar con Google Maps (red): {error}") from error

    elemento = datos[0] if isinstance(datos, list) and datos else datos
    if not isinstance(elemento, dict):
        raise ErrorGoogleMaps(f"Respuesta inesperada de Google Maps: {str(datos)[:200]}")
    if elemento.get("status", {}).get("code"):  # 0 / ausente = OK
        raise ErrorGoogleMaps(f"Google Maps: {elemento['status'].get('message', elemento['status'])}")
    if elemento.get("condition") != "ROUTE_EXISTS" or "distanceMeters" not in elemento:
        raise ErrorGoogleMaps("Google Maps no encontro una ruta en carro hasta esa ubicacion")

    segundos = int(str(elemento.get("duration", "0s")).rstrip("s") or 0)
    return {
        "km": round(elemento["distanceMeters"] / 1000, 2),
        "minutos": math.ceil(segundos / 60),
    }


def calcular_costo_envio(km, tarifa_km, redondeo=50):
    """km x tarifa, redondeado HACIA ARRIBA al multiplo de 'redondeo'.

    Ej.: 3.73 km x 500 = 1865 -> 1900.   3.7 km x 500 = 1850 -> 1850.
    """
    if tarifa_km is None or float(tarifa_km) < 0:
        raise ErrorGoogleMaps(f"Tarifa Base KM invalida: {tarifa_km!r}")
    bruto = round(float(km) * float(tarifa_km), 6)  # quita "ruido" de decimales
    if not redondeo:
        return bruto
    return int(math.ceil(bruto / redondeo) * redondeo)


def calcular_envio(destino_lat, destino_lng, tarifa_km):
    """Todo junto: ruta + zona + costo. Es la funcion que usaran bot y SaaS.

    Devuelve un diccionario:
        km, minutos, dentro_zona (True/False), radio_maximo_km,
        envio (colones, o None si esta fuera de zona),
        link_ubicacion (https://maps.google.com/?q=lat,lng)
    """
    destino_lat, destino_lng = _validar_coordenadas(destino_lat, destino_lng)
    cfg = configuracion()
    ruta = consultar_ruta(destino_lat, destino_lng)
    dentro = ruta["km"] <= cfg["radio_maximo_km"]
    return {
        "km": ruta["km"],
        "minutos": ruta["minutos"],
        "dentro_zona": dentro,
        "radio_maximo_km": cfg["radio_maximo_km"],
        "envio": calcular_costo_envio(ruta["km"], tarifa_km, cfg["redondeo"]) if dentro else None,
        "link_ubicacion": f"https://maps.google.com/?q={destino_lat},{destino_lng}",
    }


def leer_tarifa_km_del_excel():
    """Lee la Tarifa Base KM de NODO INTEGRAL C3 (fuente unica de verdad)."""
    from conexion_drive.drive_excel import leer_excel

    wb = leer_excel(solo_valores=True)
    tarifa = wb["NODO INTEGRAL"]["C3"].value
    if not isinstance(tarifa, (int, float)):
        raise ErrorGoogleMaps(f"NODO INTEGRAL C3 (Tarifa Base KM) no es un numero: {tarifa!r}")
    return tarifa


# ----------------------------------------------------------------------------
# Prueba desde la Terminal (solo LEE el Excel; no escribe nada)
# ----------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: py -m microservicios.grupo3_integraciones.google_maps <latitud> <longitud>")
        print("Ej.: py -m microservicios.grupo3_integraciones.google_maps 10.4400 -85.0900")
        sys.exit(1)

    cfg = configuracion()
    print("\n=== Configuracion ===")
    print(f"Origen (panaderia): {cfg['origen_lat']}, {cfg['origen_lng']}")
    print(f"Radio maximo: {cfg['radio_maximo_km']} km   Redondeo: {cfg['redondeo']:.0f} colones")

    print("\n=== 1) Leyendo Tarifa Base KM del Excel (NODO INTEGRAL C3) ===")
    tarifa = leer_tarifa_km_del_excel()
    print(f"OK  Tarifa Base KM: {tarifa} colones")

    print("\n=== 2) Consultando Google Maps ===")
    resultado = calcular_envio(sys.argv[1], sys.argv[2], tarifa)
    print(f"OK  Distancia en carro: {resultado['km']} km")
    print(f"    Tiempo estimado: {resultado['minutos']} min")
    print(f"    Ubicacion del cliente: {resultado['link_ubicacion']}")

    print("\n=== 3) Costo de envio ===")
    if resultado["dentro_zona"]:
        print(f"OK  Dentro de zona. Envio: {resultado['envio']:.0f} colones")
    else:
        print(f"FUERA DE ZONA: {resultado['km']} km es mas que el radio de "
              f"{resultado['radio_maximo_km']} km. El bot avisaria al cliente.")
    print("\nTODO BIEN.")
