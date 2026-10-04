"""Sesion, permisos y envio unico de formularios del SaaS de demostracion.

Vive aparte de app.py para que cada modulo (grupo1.py, ...) use los mismos permisos.
El permiso SIEMPRE se revisa en el servidor: esconder un boton no basta.
"""
import logging
import secrets
import threading
import time
from functools import wraps

from flask import abort, redirect, request, session, url_for

from . import config

log = logging.getLogger("saas_demo")

NOMBRE_ROL = {config.ROL_DESPACHO: "Despacho", config.ROL_ADMIN: "Administración"}


def es_admin():
    return session.get("rol") == config.ROL_ADMIN


def con_sesion(vista):
    @wraps(vista)
    def envoltura(*args, **kwargs):
        if session.get("rol") not in NOMBRE_ROL:
            return redirect(url_for("entrar"))
        return vista(*args, **kwargs)
    return envoltura


def solo_administracion(vista):
    @wraps(vista)
    @con_sesion
    def envoltura(*args, **kwargs):
        if not es_admin():
            log.warning("Rechazado (%s): el usuario %s no es de administracion", request.path, session.get("usuario"))
            abort(403, "Esta parte es solo para administración.")
        return vista(*args, **kwargs)
    return envoltura


# ----------------------------------------------------------------------------
# Envio unico: un formulario que REGISTRA algo nuevo (venta, compra) solo vale una vez.
# Asi un doble clic o recargar la pagina no registra la venta dos veces.
# ----------------------------------------------------------------------------
_usados = {}
_candado = threading.Lock()
_HORAS_MEMORIA = 2


def nuevo_envio():
    """Numero unico que va escondido en el formulario."""
    return secrets.token_urlsafe(12)


def envio_repetido(valor):
    """True si ese formulario ya se habia enviado. La primera vez lo anota y devuelve False."""
    if not valor or len(valor) > 40:
        abort(400, "El formulario venció. Vuelve a la página anterior, recárgala e inténtalo de nuevo.")
    ahora = time.monotonic()
    with _candado:
        for clave in [c for c, momento in _usados.items() if ahora - momento > _HORAS_MEMORIA * 3600]:
            del _usados[clave]
        if valor in _usados:
            return True
        _usados[valor] = ahora
        return False
