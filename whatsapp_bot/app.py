"""
Servidor del bot de WhatsApp (capa de conversación).

IMPORTANTE (ver Reglas_Globales_del_Proyecto, Regla #4): este servidor
NUNCA escribe en el Excel maestro directamente. Cuando un pedido queda
CONFIRMADO, solo llama a `escribir_pedido_en_excel(...)` -- una función
que hoy es un placeholder (ver el final de este archivo) y que debe
reemplazarse por la función Python COMPARTIDA con la futura web del
SaaS, cuando esa pieza se construya.

Para correr en desarrollo:
    1. pip install -r requirements.txt
    2. Configurar variables de entorno (ver .env.example)
    3. python setup_content_templates.py   (una sola vez)
    4. flask --app app run
    5. Exponer con ngrok y configurar la URL en el Sandbox de Twilio
       (ver README.md para el paso a paso completo)
"""

from flask import Flask, request
from twilio.request_validator import RequestValidator
from twilio.rest import Client

import config
from conversation import GestorConversaciones, procesar_mensaje_entrante, RespuestaBot

app = Flask(__name__)
gestor = GestorConversaciones()
twilio_client = Client(config.TWILIO_ACCOUNT_SID, config.TWILIO_AUTH_TOKEN)


def obtener_combos_activos_hoy() -> list[dict]:
    """
    Lee los combos activos HOY desde 'Tabla Combos Dinamicos' (ya
    actualizada por NODO INTEGRAL -- este bot NUNCA llama al clima ni
    a ninguna API externa directamente, ver Regla #4).

    PENDIENTE: reemplazar este placeholder por la lectura real via
    Google Drive (descargar el Excel, leer con openpyxl) -- mismo
    mecanismo que usa el robot para el resto del Excel (ver
    Reglas_Globales_del_Proyecto, Regla #1).
    """
    return []  # placeholder -- sin combos hasta conectar la lectura real


def escribir_pedido_en_excel(pedido: dict) -> str:
    """
    Escribe el pedido confirmado en 'Interface digital' y 'Detalle
    comanda digital'.

    PENDIENTE (a proposito, ver Reglas_Globales_del_Proyecto Regla #4):
    esta función debe ser la MISMA función Python compartida con la
    futura web del SaaS, no una implementación aparte aquí. Se deja
    como placeholder hasta que esa pieza compartida exista.

    Devuelve el ID_PEDIDO generado (ej. "WA-101").
    """
    raise NotImplementedError(
        "Pendiente: conectar con la función compartida de escritura a Excel "
        "(ver docstring de esta función)."
    )


@app.route("/webhook", methods=["POST"])
def webhook() -> tuple[str, int]:
    if not _validar_firma_de_twilio(request):
        return "Firma inválida", 403

    numero = request.form.get("From", "")
    texto = request.form.get("Body", "")
    boton_id = request.form.get("ButtonPayload") or request.form.get("ListId")
    lat = request.form.get("Latitude")
    lng = request.form.get("Longitude")

    ubicacion = None
    if lat and lng:
        ubicacion = (float(lat), float(lng))

    conv = gestor.obtener_o_crear(numero)

    combos_activos = obtener_combos_activos_hoy() if conv.estado.name == "ESPERANDO_UBICACION" else None

    respuesta = procesar_mensaje_entrante(
        conv,
        texto_libre=texto,
        opcion_elegida_id=boton_id,
        ubicacion=ubicacion,
        combos_activos=combos_activos,
    )

    _enviar_respuesta(numero, respuesta)

    if respuesta.tipo_mensaje == "pedido_listo_para_escribir":
        try:
            id_pedido = escribir_pedido_en_excel(respuesta.datos_extra)
            _enviar_texto(numero, f"¡Pedido confirmado! Tu número de orden es {id_pedido}. Gracias 🍞")
        except NotImplementedError:
            _enviar_texto(
                numero,
                "Tu pedido quedó armado, pero la conexión con el sistema de "
                "registro todavía no está lista. Un encargado te confirmará en breve.",
            )
        gestor.reiniciar(numero)

    return "", 204


def _validar_firma_de_twilio(req) -> bool:
    validador = RequestValidator(config.TWILIO_AUTH_TOKEN)
    firma = req.headers.get("X-Twilio-Signature", "")
    return validador.validate(req.url, req.form, firma)


def _enviar_texto(numero_destino: str, texto: str) -> None:
    twilio_client.messages.create(from_=config.TWILIO_WHATSAPP_FROM, to=numero_destino, body=texto)


def _enviar_plantilla(numero_destino: str, content_sid: str, variables: dict | None = None) -> None:
    kwargs = {"from_": config.TWILIO_WHATSAPP_FROM, "to": numero_destino, "content_sid": content_sid}
    if variables:
        import json

        kwargs["content_variables"] = json.dumps(variables)
    twilio_client.messages.create(**kwargs)


def _enviar_respuesta(numero: str, respuesta: RespuestaBot) -> None:
    tipo = respuesta.tipo_mensaje

    if tipo == "menu_principal_y_pedir_ubicacion":
        _enviar_plantilla(numero, config.CONTENT_SID_MENU_PRINCIPAL)
        _enviar_texto(numero, "Para empezar, comparte tu ubicación usando el clip 📎 → Ubicación.")

    elif tipo == "lista_combos":
        # PENDIENTE: usar una plantilla generada por el trabajo diario (ver
        # setup_content_templates.py) en vez de texto plano. Placeholder simple por ahora.
        combos = (respuesta.datos_extra or {}).get("combos_activos", [])
        texto = "Combos activos hoy:\n" + "\n".join(f"- {c}" for c in combos)
        _enviar_texto(numero, texto)

    elif tipo == "lista_productos":
        _enviar_plantilla(numero, config.CONTENT_SID_LISTA_PRODUCTOS)

    elif tipo == "pedir_cantidad":
        _enviar_texto(numero, "¿Cuántas unidades deseas?")

    elif tipo == "preguntar_agregar_mas":
        _enviar_plantilla(numero, config.CONTENT_SID_AGREGAR_MAS)

    elif tipo == "resumen_confirmacion":
        lineas = (respuesta.datos_extra or {}).get("lineas", [])
        resumen = "\n".join(f"{l['cantidad']}x {l['producto_nombre']}" for l in lineas)
        _enviar_plantilla(numero, config.CONTENT_SID_CONFIRMAR_PEDIDO, variables={"1": resumen})

    elif tipo == "cancelado":
        _enviar_texto(numero, "Pedido cancelado. Escribe cualquier cosa si quieres empezar de nuevo.")

    elif tipo == "conversacion_terminada":
        gestor.reiniciar(numero)
        _enviar_texto(numero, "Escribe cualquier cosa para empezar un nuevo pedido.")

    elif tipo == "texto":
        _enviar_texto(numero, respuesta.texto or "")


if __name__ == "__main__":
    app.run(port=5000, debug=True)
