"""Configuracion del SaaS de demostracion (Sesion 2, 3-oct-2026).

IMPORTANTE: este archivo se importa ANTES que conexion_drive y los microservicios.
El robot conoce UN solo Excel: lee EXCEL_MAESTRO_FILE_ID una vez, al arrancar. Aqui
se le entrega el ID de la COPIA de demostracion (SAAS_EXCEL_FILE_ID) en ese lugar,
asi todas las funciones compartidas (marcar_entregado, anular_pedido_digital...)
trabajan sobre la copia SIN cambiar una sola linea de su codigo.

Candados para que el SaaS nunca toque el Excel maestro:
  1. Sin SAAS_EXCEL_FILE_ID el SaaS no arranca (no usa el ID del maestro "por defecto").
  2. Si SAAS_EXCEL_FILE_ID es el mismo ID que EXCEL_MAESTRO_FILE_ID, no arranca.
  3. Antes de leer o escribir, datos.py revisa que el NOMBRE del archivo de Drive
     contenga la palabra DEMO (SAAS_PALABRA_COPIA). El maestro no la tiene.
  En Render, ademas, el servicio del SaaS no recibe el ID del maestro.

Dos modos:
  - DEMOSTRACION (normal, Render): copia de demostracion en Google Drive.
  - ENSAYO (solo en la PC, py -m saas_demo.ensayo): copia TEMPORAL de
    datos_prueba/maestro_prueba.xlsx; no toca Drive ni la copia local.

Las claves NUNCA van en el codigo ni en GitHub: viven en el .env (PC) y en las
variables de entorno de Render.
"""
import os
import re
import secrets
import sys
from pathlib import Path

from dotenv import load_dotenv

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent

ROL_DESPACHO, ROL_ADMIN = "despacho", "administracion"
LARGO_MINIMO_CLAVE = 8
# Solo para el modo ENSAYO (copia temporal en la PC). En Render las claves son obligatorias.
CLAVE_ENSAYO = "ensayo-2026"


class ConfiguracionInvalida(Exception):
    """Falta una variable o hay un valor que pondria en riesgo el Excel maestro."""


def _texto(entorno, nombre, defecto=""):
    return (entorno.get(nombre) or defecto).strip()


def _es_si(valor):
    return str(valor or "").strip().lower() in ("si", "sí", "true", "1", "yes")


def id_desde_texto(texto):
    """Acepta el ID solo o el enlace de Drive (.../d/<ID>/...). Devuelve el ID."""
    texto = (texto or "").strip().strip('"').strip("'")
    encontrado = re.search(r"/d/([A-Za-z0-9_-]{20,})", texto)
    return encontrado.group(1) if encontrado else texto


def resolver(entorno):
    """Lee y VALIDA la configuracion. No cambia nada: solo devuelve un diccionario.

    entorno: diccionario con las variables (normalmente os.environ).
    Lanza ConfiguracionInvalida con un mensaje claro si algo falta o es riesgoso.
    """
    ensayo = _es_si(entorno.get("SAAS_ENSAYO"))
    datos = {
        "ensayo": ensayo,
        "palabra_copia": _texto(entorno, "SAAS_PALABRA_COPIA", "DEMO"),
        "id_demo": None,
        "segundos_cache": int(_texto(entorno, "SAAS_SEGUNDOS_CACHE", "30")),
    }

    if ensayo:
        if _texto(entorno, "ORIGEN_EXCEL").lower() != "local":
            raise ConfiguracionInvalida(
                "El modo ENSAYO solo trabaja con una copia local. Usa:  py -m saas_demo.ensayo")
    else:
        id_demo = id_desde_texto(_texto(entorno, "SAAS_EXCEL_FILE_ID"))
        if not id_demo or "pega_aqui" in id_demo:
            raise ConfiguracionInvalida(
                "Falta SAAS_EXCEL_FILE_ID (el ID o el enlace de la COPIA de demostracion en Drive). "
                "El SaaS no arranca sin el: nunca usa el Excel maestro por defecto.")
        id_maestro = id_desde_texto(_texto(entorno, "EXCEL_MAESTRO_FILE_ID"))
        if id_maestro and id_maestro == id_demo:
            raise ConfiguracionInvalida(
                "SAAS_EXCEL_FILE_ID es el MISMO archivo que EXCEL_MAESTRO_FILE_ID. El SaaS de "
                "demostracion no puede apuntar al Excel maestro. No arranca.")
        datos["id_demo"] = id_demo

    # --- Usuarios: despacho (entrega) y administracion (entrega y anula) ---
    if ensayo:
        # En el ensayo NUNCA se usan las claves reales del .env: la pantalla de ingreso
        # del ensayo muestra estas claves de prueba para que la persona pueda entrar.
        usuario_despacho, usuario_admin = "despacho", "administracion"
        clave_despacho, clave_admin = CLAVE_ENSAYO, CLAVE_ENSAYO + "-admin"
    else:
        usuario_despacho = _texto(entorno, "SAAS_USUARIO_DESPACHO", "despacho")
        usuario_admin = _texto(entorno, "SAAS_USUARIO_ADMIN", "administracion")
        clave_despacho = _texto(entorno, "SAAS_CLAVE_DESPACHO")
        clave_admin = _texto(entorno, "SAAS_CLAVE_ADMIN")
    if usuario_despacho.lower() == usuario_admin.lower():
        raise ConfiguracionInvalida("SAAS_USUARIO_DESPACHO y SAAS_USUARIO_ADMIN deben ser nombres distintos.")
    for nombre, clave in (("SAAS_CLAVE_DESPACHO", clave_despacho), ("SAAS_CLAVE_ADMIN", clave_admin)):
        if not clave:
            raise ConfiguracionInvalida(f"Falta {nombre} (la clave se pone en el .env o en Render, nunca en el codigo).")
        if len(clave) < LARGO_MINIMO_CLAVE:
            raise ConfiguracionInvalida(f"{nombre} debe tener al menos {LARGO_MINIMO_CLAVE} caracteres.")
    if clave_despacho == clave_admin:
        raise ConfiguracionInvalida("SAAS_CLAVE_DESPACHO y SAAS_CLAVE_ADMIN deben ser distintas.")
    datos["usuarios"] = {
        usuario_despacho.lower(): {"usuario": usuario_despacho, "clave": clave_despacho, "rol": ROL_DESPACHO},
        usuario_admin.lower(): {"usuario": usuario_admin, "clave": clave_admin, "rol": ROL_ADMIN},
    }

    # Llave que firma la cookie de la sesion. Si no se pone, se crea una al arrancar
    # (al reiniciar el servicio las personas vuelven a ingresar; no hay otro efecto).
    datos["llave_sesion"] = _texto(entorno, "SAAS_SECRET_KEY") or secrets.token_hex(32)
    return datos


# ----------------------------------------------------------------------------
# Al importar: leer el .env, validar y PREPARAR el entorno para el robot
# ----------------------------------------------------------------------------
load_dotenv(RAIZ_PROYECTO / ".env")          # no sobrescribe variables que ya existan
_datos = resolver(os.environ)

ENSAYO = _datos["ensayo"]
ID_DEMO = _datos["id_demo"]
PALABRA_COPIA = _datos["palabra_copia"]
SEGUNDOS_CACHE = _datos["segundos_cache"]
USUARIOS = _datos["usuarios"]
LLAVE_SESION = _datos["llave_sesion"]
HORAS_SESION = 8
INTENTOS_ANTES_DE_BLOQUEO = 5
MINUTOS_BLOQUEO = 5

if not ENSAYO:
    # El robot leera ESTE ID como "su" Excel. En Render el ID del maestro ni siquiera existe.
    os.environ["ORIGEN_EXCEL"] = "drive"
    os.environ["EXCEL_MAESTRO_FILE_ID"] = ID_DEMO
    _config_robot = sys.modules.get("conexion_drive.config")
    if _config_robot is not None:              # por si algo lo importo antes (no deberia)
        _config_robot.ORIGEN_EXCEL = "drive"
        _config_robot.EXCEL_MAESTRO_FILE_ID = ID_DEMO
