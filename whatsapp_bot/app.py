"""Servidor del bot de WhatsApp para Twilio (version 29-sep-2026, Sesion A del bot).

Flujo de cada mensaje:
  1. Twilio llama a /webhook con el mensaje (texto, ubicacion, nombre, telefono,
     MessageSid). Se revisa la firma: solo Twilio puede usar este webhook.
  2. El nucleo decide la respuesta; se devuelve AL INSTANTE dentro de la misma
     llamada (TwiML), asi el cliente no espera.
  3. Si hay que escribir en el Excel (pedido confirmado, entrega), se pone en la
     fila de escrituras; al terminar, el resultado se envia con la API de Twilio.

Correr (desde la carpeta ROBOT PANADERO):
    py -m whatsapp_bot.app                         (desarrollo, puerto 5000)
En Render (fase de hosting) se corre con UN SOLO proceso, porque las
conversaciones viven en memoria:
    gunicorn --workers 1 --threads 4 whatsapp_bot.app:app
"""
import logging

from flask import Flask, request

from . import config
from .nucleo import Bot

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("whatsapp_bot")

app = Flask(__name__)
_cliente_twilio = None


def _twilio():
    global _cliente_twilio
    if _cliente_twilio is None:
        from twilio.rest import Client
        _cliente_twilio = Client(config.TWILIO_ACCOUNT_SID, config.TWILIO_AUTH_TOKEN)
    return _cliente_twilio


def enviar_por_twilio(numero_destino, texto):
    """Mensajes FUERA de la respuesta inmediata (confirmacion, recibo, aviso al repartidor)."""
    try:
        _twilio().messages.create(from_=config.TWILIO_WHATSAPP_FROM, to=numero_destino, body=texto)
    except Exception:
        log.exception("No se pudo enviar el mensaje por Twilio")


bot = Bot(enviar=enviar_por_twilio)


def _firma_valida(req):
    if not config.VALIDAR_FIRMA_TWILIO:
        return True
    from twilio.request_validator import RequestValidator
    url = req.url
    if req.headers.get("X-Forwarded-Proto") == "https" and url.startswith("http://"):
        url = "https://" + url[len("http://"):]          # detras del proxy de Render la URL publica es https
    return RequestValidator(config.TWILIO_AUTH_TOKEN).validate(url, req.form, req.headers.get("X-Twilio-Signature", ""))


def _twiml(texto):
    """Respuesta inmediata en formato TwiML (el XML que Twilio entiende)."""
    from xml.sax.saxutils import escape
    cuerpo = f"<Message>{escape(texto)}</Message>" if texto else ""
    xml = f'<?xml version="1.0" encoding="UTF-8"?><Response>{cuerpo}</Response>'
    return xml, 200, {"Content-Type": "application/xml; charset=utf-8"}


@app.route("/webhook", methods=["POST"])
def webhook():
    if not _firma_valida(request):
        log.warning("Firma de Twilio invalida: mensaje rechazado")
        return "Firma invalida", 403
    lat, lng = request.form.get("Latitude"), request.form.get("Longitude")
    ubicacion = (float(lat), float(lng)) if lat and lng else None
    texto, accion = bot.recibir(
        numero=request.form.get("From", ""),
        texto=request.form.get("Body", ""),
        nombre=request.form.get("ProfileName"),
        id_mensaje=request.form.get("MessageSid"),
        ubicacion=ubicacion,
    )
    bot.programar(accion)           # la escritura sigue en segundo plano
    return _twiml(texto)


@app.route("/salud", methods=["GET"])
def salud():
    """Para revisar que el servidor esta vivo (lo usara Render)."""
    return {"estado": "ok", "escrituras_pendientes": bot.fila.pendientes()}


if __name__ == "__main__":
    app.run(port=5000, debug=False)
