"""
Configuración central del bot de WhatsApp.

Todo se lee de variables de entorno -- nunca se escriben credenciales
directo en el código. Para desarrollo local, crea un archivo .env
(ver .env.example) y cárgalo con `python-dotenv` o exportando las
variables en tu terminal antes de correr app.py.
"""

import os

# --- Credenciales de Twilio (obligatorias) ---
TWILIO_ACCOUNT_SID = os.environ.get("TWILIO_ACCOUNT_SID", "")
TWILIO_AUTH_TOKEN = os.environ.get("TWILIO_AUTH_TOKEN", "")

# Número de WhatsApp del negocio, formato "whatsapp:+506XXXXXXXX".
# Mientras se usa el Sandbox de pruebas de Twilio, este es el número
# de sandbox que Twilio te asigna (ver README.md).
TWILIO_WHATSAPP_FROM = os.environ.get("TWILIO_WHATSAPP_FROM", "")

# --- IDs de las plantillas de contenido (Content Templates) ---
# Estos IDs (empiezan con "HX...") se generan UNA SOLA VEZ corriendo
# setup_content_templates.py, y de ahí se copian aquí o a las
# variables de entorno reales del servidor.
CONTENT_SID_MENU_PRINCIPAL = os.environ.get("CONTENT_SID_MENU_PRINCIPAL", "")
CONTENT_SID_LISTA_PRODUCTOS = os.environ.get("CONTENT_SID_LISTA_PRODUCTOS", "")
CONTENT_SID_LISTA_COMBOS = os.environ.get("CONTENT_SID_LISTA_COMBOS", "")
CONTENT_SID_CONFIRMAR_PEDIDO = os.environ.get("CONTENT_SID_CONFIRMAR_PEDIDO", "")
CONTENT_SID_AGREGAR_MAS = os.environ.get("CONTENT_SID_AGREGAR_MAS", "")

# --- Catálogo de productos propios (Baguette y Cupcakes) ---
# NOTA IMPORTANTE (ver Reglas_Globales_del_Proyecto, Regla #4): esto es
# solo el catálogo para MOSTRAR el menú al cliente. El precio final
# que se escribe en el pedido SIEMPRE debe leerse en el momento real
# desde el Excel maestro (via Microsoft Graph API) -- nunca de aquí
# directamente -- porque el precio puede cambiar. Esta lista es solo
# para armar la plantilla de WhatsApp (id, nombre visible, descripción).
PRODUCTOS_PROPIOS = [
    {"id": "PROD_BAGUETTE", "titulo": "Pan Baguette", "descripcion": "Pan artesanal recién horneado"},
    {"id": "PROD_CUPCAKES", "titulo": "Cupcakes", "descripcion": "Cupcakes decorados, unidad"},
]

# Catálogo de productos de terceros (Musmanni-style: reventa de otras marcas)
PRODUCTOS_TERCEROS = [
    {"id": "PROD_LECHE", "titulo": "Leche", "descripcion": "Leche Dos Pinos semidescremada 1L"},
    {"id": "PROD_QUESO", "titulo": "Queso", "descripcion": "Queso Zarcero fresco semiduro"},
    {"id": "PROD_MARGARINA", "titulo": "Margarina", "descripcion": "Margarina Numar"},
    {"id": "PROD_COCACOLA", "titulo": "Coca Cola", "descripcion": "Refresco 600ml"},
    {"id": "PROD_NATILLA", "titulo": "Natilla", "descripcion": "Natilla Dos Pinos"},
]

# Tiempo, en segundos, que se mantiene abierta una conversación sin
# actividad antes de considerarla "abandonada" y limpiar su estado.
CONVERSACION_TIMEOUT_SEGUNDOS = 30 * 60  # 30 minutos
