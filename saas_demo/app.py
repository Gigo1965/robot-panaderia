"""Servidor web del SaaS de demostracion (Sesion 2, 3-oct-2026; Decisiones Pendientes #54).

Que es: unas pocas pantallas para que la panaderia (y los tutores) vean los pedidos
digitales y hagan dos cosas sobre la COPIA de demostracion del Excel:
  - "Marcar entregado"  -> marcar_entregado()        (despacho y administracion)
  - "Anular pedido"     -> anular_pedido_digital()   (SOLO administracion)
Son las MISMAS funciones que usa el bot de WhatsApp (Regla #5): aqui no hay reglas
de negocio, solo pantallas, usuarios y permisos.

Diferencia con el bot: el SaaS no envia WhatsApp. El recibo de la entrega se
muestra en pantalla (trabaja sobre la copia; no tiene las claves de Twilio).

Correr:
    En la PC, ensayo sobre una copia temporal:   py -m saas_demo.ensayo
    En Render (servicio aparte del bot):
        gunicorn --workers 1 --threads 4 --timeout 120 saas_demo.app:app
    (1 solo proceso: el candado de escrituras y el cache viven en memoria.)
"""
import hmac
import logging
import secrets
import time
from datetime import date, datetime, time as hora_del_dia, timedelta
from functools import wraps

from flask import Flask, abort, flash, redirect, render_template, request, session, url_for

from . import config                       # PRIMERO (prepara el entorno del robot)
from . import datos                        # noqa: E402
from conexion_drive import ConflictoDeVersion, ErrorConexionExcel                     # noqa: E402
from microservicios.grupo1_captura_operativa import pedidos_digitales as pd          # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("saas_demo")

app = Flask(__name__, template_folder="plantillas", static_folder="estaticos", static_url_path="/estaticos")
app.secret_key = config.LLAVE_SESION
app.config.update(
    SESSION_COOKIE_NAME="saas_demo",
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=not config.ENSAYO,          # en Render solo viaja por https
    PERMANENT_SESSION_LIFETIME=timedelta(hours=config.HORAS_SESION),
    MAX_CONTENT_LENGTH=64 * 1024,
)

NOMBRE_ROL = {config.ROL_DESPACHO: "Despacho", config.ROL_ADMIN: "Administración"}
LARGO_MAXIMO_MOTIVO = 200
_intentos = {}          # direccion -> [fallos, bloqueado_hasta]


# ----------------------------------------------------------------------------
# Como se muestran los datos
# ----------------------------------------------------------------------------
MESES = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")


@app.template_filter("colones")
def colones(valor):
    try:
        return f"₡{float(valor):,.2f}"
    except (TypeError, ValueError):
        return ""


@app.template_filter("fecha")
def fecha(valor):
    if isinstance(valor, (datetime, date)):
        return f"{valor.day} {MESES[valor.month - 1]} {valor.year}"
    return str(valor or "")


@app.template_filter("hora")
def hora(valor):
    if isinstance(valor, (datetime, hora_del_dia)):
        h12 = valor.hour % 12 or 12
        return f"{h12}:{valor.minute:02d} {'a. m.' if valor.hour < 12 else 'p. m.'}"
    return str(valor or "")


@app.template_filter("fecha_hora")
def fecha_hora(valor):
    return f"{fecha(valor)}, {hora(valor)}" if isinstance(valor, datetime) else str(valor or "")


@app.context_processor
def _para_todas_las_pantallas():
    return {
        "rol": session.get("rol"),
        "nombre_rol": NOMBRE_ROL.get(session.get("rol"), ""),
        "usuario": session.get("usuario"),
        "es_admin": session.get("rol") == config.ROL_ADMIN,
        "ficha": session.get("ficha", ""),
        "ensayo": config.ENSAYO,
        # Solo en el ENSAYO de la PC (copia temporal) se muestran las claves de prueba.
        "cuentas_ensayo": list(config.USUARIOS.values()) if config.ENSAYO else [],
    }


# ----------------------------------------------------------------------------
# Seguridad: sesion, ficha de formularios y permisos
# ----------------------------------------------------------------------------
@app.before_request
def _antes():
    if "ficha" not in session:                      # ficha contra formularios falsos (CSRF)
        session["ficha"] = secrets.token_urlsafe(24)
    if request.method == "POST":
        enviada = request.form.get("ficha", "")
        if not enviada or not hmac.compare_digest(enviada, session["ficha"]):
            abort(400, "El formulario venció. Vuelve a la página anterior, recárgala e inténtalo de nuevo.")


@app.after_request
def _cabeceras(respuesta):
    respuesta.headers["X-Frame-Options"] = "DENY"
    respuesta.headers["X-Content-Type-Options"] = "nosniff"
    respuesta.headers["Referrer-Policy"] = "same-origin"
    respuesta.headers["Content-Security-Policy"] = "default-src 'self'; frame-ancestors 'none'; form-action 'self'"
    if not request.path.startswith("/estaticos/"):
        respuesta.headers["Cache-Control"] = "no-store"
    return respuesta


def con_sesion(vista):
    @wraps(vista)
    def envoltura(*args, **kwargs):
        if session.get("rol") not in NOMBRE_ROL:
            return redirect(url_for("entrar"))
        return vista(*args, **kwargs)
    return envoltura


def solo_administracion(vista):
    """El permiso se revisa en el SERVIDOR: esconder el boton no basta (#54)."""
    @wraps(vista)
    @con_sesion
    def envoltura(*args, **kwargs):
        if session.get("rol") != config.ROL_ADMIN:
            log.warning("Anulacion rechazada: el usuario %s no es de administracion", session.get("usuario"))
            abort(403, "Solo administración puede anular un pedido.")
        return vista(*args, **kwargs)
    return envoltura


def _direccion():
    return (request.headers.get("X-Forwarded-For", "").split(",")[0].strip() or request.remote_addr or "?")


def _bloqueado(direccion):
    fallos, hasta = _intentos.get(direccion, (0, 0.0))
    return fallos >= config.INTENTOS_ANTES_DE_BLOQUEO and time.monotonic() < hasta


def _anotar_fallo(direccion):
    fallos, hasta = _intentos.get(direccion, (0, 0.0))
    if fallos >= config.INTENTOS_ANTES_DE_BLOQUEO and time.monotonic() >= hasta:
        fallos = 0                                   # ya cumplio el bloqueo: empieza de nuevo
    _intentos[direccion] = (fallos + 1, time.monotonic() + config.MINUTOS_BLOQUEO * 60)


# ----------------------------------------------------------------------------
# Ingreso
# ----------------------------------------------------------------------------
@app.route("/entrar", methods=["GET", "POST"])
def entrar():
    if session.get("rol") in NOMBRE_ROL:
        return redirect(url_for("pedidos"))
    if request.method == "POST":
        direccion = _direccion()
        if _bloqueado(direccion):
            flash(f"Demasiados intentos. Espera {config.MINUTOS_BLOQUEO} minutos y vuelve a probar.", "error")
            return render_template("entrar.html"), 429
        nombre = request.form.get("usuario", "").strip().lower()
        clave = request.form.get("clave", "")
        cuenta = config.USUARIOS.get(nombre)
        clave_buena = cuenta["clave"] if cuenta else secrets.token_hex(8)
        coincide = hmac.compare_digest(clave.encode("utf-8"), clave_buena.encode("utf-8"))
        if cuenta and coincide:
            _intentos.pop(direccion, None)
            ficha = session.get("ficha")
            session.clear()
            session.permanent = True
            session.update(usuario=cuenta["usuario"], rol=cuenta["rol"], ficha=ficha)
            log.info("Ingreso: %s (%s)", cuenta["usuario"], cuenta["rol"])
            return redirect(url_for("pedidos"))
        _anotar_fallo(direccion)
        time.sleep(0.5)
        log.warning("Ingreso fallido desde %s", direccion)
        flash("El usuario o la contraseña no coinciden.", "error")
        return render_template("entrar.html"), 401
    return render_template("entrar.html")


@app.route("/salir", methods=["POST"])
def salir():
    session.clear()
    return redirect(url_for("entrar"))


# ----------------------------------------------------------------------------
# Pedidos
# ----------------------------------------------------------------------------
FILTROS = {
    "por-entregar": ("Por entregar", datos.PENDIENTE),
    "entregados": ("Entregados", datos.ENTREGADO),
    "anulados": ("Anulados", datos.ANULADO),
    "todos": ("Todos", None),
}


def _leer(forzar=False):
    """Lee los pedidos; si el Excel no se puede leer, devuelve (None, mensaje)."""
    try:
        return datos.leer_pedidos(forzar), None
    except datos.CopiaNoVerificada as error:
        log.error("Copia no verificada: %s", error)
        return None, str(error)
    except (ErrorConexionExcel, KeyError) as error:
        log.exception("No se pudo leer el Excel")
        return None, f"No se pudo leer el Excel de demostración: {error}"


@app.route("/")
@con_sesion
def pedidos():
    filtro = request.args.get("ver", "por-entregar")
    if filtro not in FILTROS:
        filtro = "por-entregar"
    libro, problema = _leer(forzar=request.args.get("actualizar") == "1")
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    situacion = FILTROS[filtro][1]
    lista = [p for p in libro["pedidos"] if situacion is None or p["situacion"] == situacion]
    pestanas = [(clave, titulo, libro["cuentas"]["todos" if s is None else s]) for clave, (titulo, s) in FILTROS.items()]
    return render_template("pedidos.html", libro=libro, lista=lista, filtro=filtro, pestanas=pestanas)


@app.route("/pedidos/<id_pedido>")
@con_sesion
def pedido(id_pedido):
    libro, problema = _leer()
    if libro is None:
        return render_template("problema.html", problema=problema), 503
    encontrado = datos.buscar_pedido(id_pedido)
    if encontrado is None:
        abort(404, f"No existe el pedido {id_pedido} en el Excel de demostración.")
    return render_template("pedido.html", libro=libro, p=encontrado, largo_motivo=LARGO_MAXIMO_MOTIVO)


def _pedido_o_404(id_pedido):
    try:
        encontrado = datos.buscar_pedido(id_pedido)
    except (datos.CopiaNoVerificada, ErrorConexionExcel) as error:
        abort(503, str(error))
    if encontrado is None:
        abort(404, f"No existe el pedido {id_pedido} en el Excel de demostración.")
    return encontrado


def _avisar_rechazo(error):
    """Los rechazos vienen de la funcion compartida: se muestran tal cual los escribio el robot."""
    if isinstance(error, (pd.AnulacionNoPermitida, pd.EntregaNoPermitida, LookupError, ValueError)):
        flash(f"El robot no lo permitió: {error}", "error")
    elif isinstance(error, datos.Ocupado):
        flash(str(error), "error")
    elif isinstance(error, datos.CopiaNoVerificada):
        log.error("Copia no verificada: %s", error)
        flash(str(error), "error")
    elif isinstance(error, ConflictoDeVersion):
        flash("El Excel cambió mientras se guardaba y no se pudo terminar. No se escribió nada; inténtalo de nuevo.",
              "error")
    else:
        log.exception("Fallo una escritura del SaaS")
        flash(f"No se pudo guardar en el Excel. No se escribió nada. Detalle: {error}", "error")


@app.route("/pedidos/<id_pedido>/entregar", methods=["POST"])
@con_sesion
def entregar(id_pedido):
    encontrado = _pedido_o_404(id_pedido)
    try:
        resultado, segundos = datos.entregar(encontrado["id"], session["rol"])
    except Exception as error:          # noqa: BLE001 (cada caso se explica en _avisar_rechazo)
        _avisar_rechazo(error)
        return redirect(url_for("pedido", id_pedido=encontrado["id"]))
    if resultado.get("estado") == "YA_ENTREGADO":
        flash(f"{encontrado['id']} ya estaba entregado. No se escribió nada.", "aviso")
    else:
        flash(f"{encontrado['id']} quedó entregado. Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("pedido", id_pedido=encontrado["id"]))


@app.route("/pedidos/<id_pedido>/anular", methods=["POST"])
@solo_administracion
def anular(id_pedido):
    encontrado = _pedido_o_404(id_pedido)
    # El motivo nunca queda al inicio de una celda: el robot lo escribe como
    # "ANULADO: <motivo> (anula WA-xxx)", asi Excel no puede leerlo como formula.
    motivo = " ".join(request.form.get("motivo", "").split())[:LARGO_MAXIMO_MOTIVO]
    try:
        resultado, segundos = datos.anular(encontrado["id"], motivo, session["rol"])
    except Exception as error:          # noqa: BLE001
        _avisar_rechazo(error)
        return redirect(url_for("pedido", id_pedido=encontrado["id"]))
    flash(f"{encontrado['id']} quedó anulado ({resultado['id_anulacion']}, "
          f"{colones(resultado['total_revertido'])}). Guardado en el Excel en {segundos} s.", "ok")
    return redirect(url_for("pedido", id_pedido=encontrado["id"]))


# ----------------------------------------------------------------------------
# Otros
# ----------------------------------------------------------------------------
@app.route("/salud")
def salud():
    """Para Render: dice que el servidor esta vivo. No lee el Excel."""
    return {"estado": "ok", "servicio": "saas_demo", "modo": "ensayo" if config.ENSAYO else "demostracion"}


@app.route("/robots.txt")
def robots():
    return "User-agent: *\nDisallow: /\n", 200, {"Content-Type": "text/plain; charset=utf-8"}


@app.errorhandler(400)
@app.errorhandler(403)
@app.errorhandler(404)
@app.errorhandler(503)
def _error(error):
    return render_template("problema.html", problema=error.description, codigo=error.code), error.code


if __name__ == "__main__":
    app.run(port=5001, debug=False)
