"""OBSOLETO (29-sep-2026): el bot ya NO usa plantillas de Twilio. Decision del usuario:
MENUS NUMERADOS leidos del Excel en el momento (el catalogo cambia cada dia y las
plantillas son fijas). Se conserva solo como historia; no hace falta correrlo.


Script de configuración inicial -- se corre UNA sola vez (o cada vez
que cambie el catálogo de productos) para crear las plantillas de
contenido (Content Templates) que Twilio necesita para enviar listas
y botones interactivos por WhatsApp.

Uso:
    python setup_content_templates.py

Al terminar, copia los ContentSid (empiezan con "HX...") que imprime
a tus variables de entorno (o a config.py) antes de correr app.py.

Por qué existe este paso: a diferencia de la API cruda de Meta, Twilio
no permite mandar un mensaje de lista o de botones "al vuelo" -- exige
crear la plantilla primero. Dentro de las 24 horas después de que el
cliente escribe, se puede enviar sin que WhatsApp la apruebe primero
(ideal para probar en el Sandbox); para producción a largo plazo, se
recomienda de igual forma someterla a aprobación de WhatsApp.
"""

from twilio.rest import Client

import config


def crear_plantilla_menu_principal(client: Client) -> str:
    respuesta = client.content.v1.contents.create(
        friendly_name="menu_principal_panaderia",
        language="es",
        types={
            "twilio/quick-reply": {
                "body": "¡Hola! 👋 Bienvenido a la panadería. ¿Qué deseas hacer?",
                "actions": [
                    {"title": "Hacer pedido", "id": "HACER_PEDIDO"},
                    {"title": "Ver combos de hoy", "id": "VER_COMBOS"},
                    {"title": "Hablar con alguien", "id": "HABLAR_CON_ALGUIEN"},
                ],
            }
        },
    )
    return respuesta.sid


def crear_plantilla_lista_productos(client: Client) -> str:
    filas = []
    for p in config.PRODUCTOS_PROPIOS + config.PRODUCTOS_TERCEROS:
        filas.append({"item": p["titulo"], "id": p["id"], "description": p["descripcion"]})

    respuesta = client.content.v1.contents.create(
        friendly_name="lista_productos_panaderia",
        language="es",
        types={
            "twilio/list-picker": {
                "body": "Elige un producto:",
                "button": "Ver productos",
                "items": filas,
            }
        },
    )
    return respuesta.sid


def crear_plantilla_agregar_mas(client: Client) -> str:
    respuesta = client.content.v1.contents.create(
        friendly_name="agregar_mas_productos",
        language="es",
        types={
            "twilio/quick-reply": {
                "body": "Producto agregado. ¿Deseas agregar otro?",
                "actions": [
                    {"title": "Agregar otro", "id": "AGREGAR_OTRO"},
                    {"title": "Ya terminé", "id": "TERMINAR_PEDIDO"},
                ],
            }
        },
    )
    return respuesta.sid


def crear_plantilla_confirmar_pedido(client: Client) -> str:
    respuesta = client.content.v1.contents.create(
        friendly_name="confirmar_pedido_panaderia",
        language="es",
        types={
            "twilio/quick-reply": {
                "body": "Este es tu pedido:\n{{1}}\n\n¿Confirmas?",
                "actions": [
                    {"title": "Confirmar", "id": "CONFIRMAR"},
                    {"title": "Cancelar", "id": "CANCELAR"},
                ],
            }
        },
    )
    return respuesta.sid


def main() -> None:
    if not config.TWILIO_ACCOUNT_SID or not config.TWILIO_AUTH_TOKEN:
        raise SystemExit(
            "Faltan TWILIO_ACCOUNT_SID / TWILIO_AUTH_TOKEN. "
            "Configura las variables de entorno antes de correr este script."
        )

    client = Client(config.TWILIO_ACCOUNT_SID, config.TWILIO_AUTH_TOKEN)

    print("Creando plantillas de contenido en Twilio...\n")

    sid_menu = crear_plantilla_menu_principal(client)
    print(f"CONTENT_SID_MENU_PRINCIPAL={sid_menu}")

    sid_productos = crear_plantilla_lista_productos(client)
    print(f"CONTENT_SID_LISTA_PRODUCTOS={sid_productos}")

    sid_agregar_mas = crear_plantilla_agregar_mas(client)
    print(f"CONTENT_SID_AGREGAR_MAS={sid_agregar_mas}")

    sid_confirmar = crear_plantilla_confirmar_pedido(client)
    print(f"CONTENT_SID_CONFIRMAR_PEDIDO={sid_confirmar}")

    print(
        "\nCopia estas 4 líneas a tu archivo .env (o a las variables de entorno "
        "del servidor) antes de correr app.py.\n"
        "NOTA: la lista de COMBOS se genera dinámicamente cada vez que hay un "
        "combo activo (via 'Tabla Combos Dinamicos'), así que no tiene una "
        "plantilla fija aquí -- se construye en tiempo real en app.py usando "
        "el mismo tipo 'twilio/list-picker'."
    )


if __name__ == "__main__":
    main()
