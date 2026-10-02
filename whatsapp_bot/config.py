"""Configuracion del bot de WhatsApp (version 29-sep-2026, Sesion A del bot).

Todo lo secreto o personal se lee del archivo .env de la RAIZ del proyecto
(ROBOT PANADERO\\.env), igual que conexion_drive. Nunca se escriben claves ni
telefonos reales en el codigo ni en GitHub.

Los datos del NEGOCIO (productos, precios, combos, SINPE, promocion del dia)
NO viven aqui: el bot los lee del Excel maestro en el momento (Regla #4, #16).
"""
import os
from pathlib import Path

from dotenv import load_dotenv

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
load_dotenv(RAIZ_PROYECTO / ".env")          # no sobrescribe variables que ya existan


def _leer(nombre, defecto=""):
    return (os.environ.get(nombre) or defecto).strip()


# --- Credenciales de Twilio (en el .env; nunca en el chat ni en GitHub) ---
TWILIO_ACCOUNT_SID = _leer("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = _leer("TWILIO_AUTH_TOKEN")
# Numero del Sandbox de Twilio, formato "whatsapp:+14155238886" (decision 29-sep-2026: Sandbox)
TWILIO_WHATSAPP_FROM = _leer("TWILIO_WHATSAPP_FROM")

# --- Numeros autorizados para el comando "ENTREGADO WA-xxx" (opcion C, 29-sep-2026) ---
# En el .env, separados por coma:  BOT_NUMEROS_AUTORIZADOS=+506XXXXXXXX,+506YYYYYYYY
# Van en el .env (no en el Excel) para que no aparezcan en la copia de demostracion.
NUMEROS_AUTORIZADOS_TEXTO = _leer("BOT_NUMEROS_AUTORIZADOS")

# --- Reglas propias del bot (no son datos del negocio) ---
MINUTOS_INACTIVIDAD = int(_leer("BOT_MINUTOS_INACTIVIDAD", "30"))   # decision del usuario: 30 min
MINUTOS_ENTREGA_ESTIMADOS = int(_leer("BOT_MINUTOS_ENTREGA", "25"))  # texto #28 aprobado
CANTIDAD_MAXIMA_POR_LINEA = 99
SEGUNDOS_CACHE_CATALOGO = 120        # el catalogo se relee del Excel cada 2 minutos como maximo

# Revisar que cada mensaje venga de verdad de Twilio (firma X-Twilio-Signature).
VALIDAR_FIRMA_TWILIO = _leer("BOT_VALIDAR_FIRMA", "si").lower() in ("si", "sí", "true", "1")
# || V15 (1-oct-2026): direccion PUBLICA exacta del webhook, la misma que esta en el
# Sandbox de Twilio (ej. https://robot-panaderia.onrender.com/webhook). Twilio firma
# con esa direccion; detras del proxy de Render el bot puede ver otra. Opcional:
# si no se pone, el bot arma la direccion desde la peticion (como antes).
URL_PUBLICA_WEBHOOK = _leer("BOT_URL_PUBLICA")

# ---------------------------------------------------------------------------
# OBSOLETO (29-sep-2026): el bot ya no usa plantillas de Twilio ni un catalogo
# escrito a mano. Decision del usuario: MENUS NUMERADOS, porque el catalogo
# cambia cada dia (combos por clima, feriados) y las plantillas son fijas.
# Se conservan los nombres solo como historia; nada los lee.
#   CONTENT_SID_MENU_PRINCIPAL, CONTENT_SID_LISTA_PRODUCTOS, CONTENT_SID_LISTA_COMBOS,
#   CONTENT_SID_CONFIRMAR_PEDIDO, CONTENT_SID_AGREGAR_MAS (ver setup_content_templates.py)
#   PRODUCTOS_PROPIOS / PRODUCTOS_TERCEROS con IDs PROD_BAGUETTE... (no coincidian con la
#   TABLA CRUCE MAESTRA, Regla #26).
# ---------------------------------------------------------------------------
