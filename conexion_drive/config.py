"""Lee la configuracion del robot desde el archivo .env (o variables de entorno).

En tu PC los valores vienen del archivo .env de la raiz del proyecto.
En la nube (GitHub Actions / Render) vienen de los "Secrets" / "Environment".
"""
import os
from pathlib import Path

from dotenv import load_dotenv

# Carpeta raiz del proyecto: ROBOT_PANADERIA/
RAIZ_PROYECTO = Path(__file__).resolve().parent.parent

# load_dotenv NO sobrescribe variables que ya existan en el sistema
load_dotenv(RAIZ_PROYECTO / ".env")


def _leer(nombre, defecto=None):
    valor = os.getenv(nombre, defecto)
    if isinstance(valor, str):
        valor = valor.strip()
        return valor or defecto
    return valor


def _ruta(valor):
    """Convierte una ruta relativa (ej. credenciales/x.json) en absoluta desde la raiz."""
    if not valor:
        return None
    ruta = Path(valor)
    return ruta if ruta.is_absolute() else RAIZ_PROYECTO / ruta


ORIGEN_EXCEL = (_leer("ORIGEN_EXCEL", "drive") or "drive").lower()
EXCEL_MAESTRO_FILE_ID = _leer("EXCEL_MAESTRO_FILE_ID")
GOOGLE_SERVICE_ACCOUNT_FILE = _ruta(_leer("GOOGLE_SERVICE_ACCOUNT_FILE"))
GOOGLE_SERVICE_ACCOUNT_JSON = _leer("GOOGLE_SERVICE_ACCOUNT_JSON")
EXCEL_LOCAL_PATH = _ruta(_leer("EXCEL_LOCAL_PATH"))
SOFFICE_PATH = _leer("SOFFICE_PATH")
RECALCULAR = (_leer("RECALCULAR", "si") or "si").lower() in ("si", "sí", "true", "1", "yes")
ZONA_HORARIA = _leer("ZONA_HORARIA", "America/Costa_Rica")
