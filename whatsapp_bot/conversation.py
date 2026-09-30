"""Conversacion del bot de WhatsApp (version 29-sep-2026, Sesion A del bot).

Decisiones del usuario (29-sep-2026):
  - Bot SOLO EN ESPANOL. MENUS NUMERADOS ("responde con el numero"): el catalogo
    se lee del Excel en el momento (combos por clima, feriados), sin plantillas.
  - Pedido sin confirmar: se borra solo despues de 30 minutos SIN mensajes del
    cliente (el reloj se reinicia con cada mensaje). Regla automatica: el bot
    decide solo sobre lo que todavia no toco el Excel; las personas deciden sobre
    lo registrado (anular = administracion, entregar = repartidor).
  - Mensaje de confirmacion #28 aprobado; forma de pago obligatoria (Regla #31);
    lineas se quitan solo ANTES de confirmar (Regla #30); envio o retiro.
  - Comando "ENTREGADO WA-120" solo desde numeros autorizados (.env).
  - Sin opcion "Hablar con alguien" (el bot no lleva a ninguna persona): el saludo
    muestra directamente el menu de productos. Para consultas o errores, el
    mensaje de pedido registrado y el recibo dicen que se comunique con la
    panaderia (atencion humana = mejora para produccion).

Este archivo NO habla con Twilio ni escribe en el Excel: solo decide que
responder. Lo que necesita leer (catalogo, cotizaciones) lo pide a "servicios"
(conector_excel.ConectorExcel). Las escrituras las devuelve como una ACCION
("registrar" / "entregar") que app.py o el simulador ponen en la fila.
"""
import logging
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from enum import Enum, auto
from threading import RLock

log = logging.getLogger("whatsapp_bot")

CANTIDAD_MAXIMA = 99
MESES = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
AVISO_28 = ("Importante: al confirmar, tu pedido queda en firme. Por higiene, los alimentos que salen de "
            "la panadería no se reciben de vuelta. Guardamos tu número solo para avisarte sobre tu pedido.")
TEXTO_CANCELADO = "Pedido cancelado. Escribe cualquier cosa si quieres empezar de nuevo."
TEXTO_AYUDA = "¿Algún problema con tu pedido? Comunícate directamente con la panadería."


class ErrorParaCliente(Exception):
    """Error del negocio con un mensaje apto para el cliente (lo lanza el conector)."""


class TotalCambiado(ErrorParaCliente):
    """El total cambio entre el resumen y el 'Si' (ej. se apago un combo)."""


class Estado(Enum):
    INICIO = auto()
    ELIGIENDO_PRODUCTO = auto()
    CANTIDAD = auto()
    ALGO_MAS = auto()
    QUITANDO = auto()
    ENTREGA = auto()
    UBICACION = auto()
    FUERA_ZONA = auto()
    PAGO = auto()
    CONFIRMAR = auto()
    REGISTRANDO = auto()


@dataclass
class MensajeEntrante:
    numero: str                    # "whatsapp:+506..." tal como lo manda Twilio
    texto: str = ""
    nombre: str | None = None      # nombre de perfil de WhatsApp (ProfileName)
    id_mensaje: str | None = None  # MessageSid (evita pedidos duplicados)
    ubicacion: tuple | None = None # (latitud, longitud) si mando el pin
    autorizado: bool = False       # numero de la lista BOT_NUMEROS_AUTORIZADOS


@dataclass
class Respuesta:
    mensajes: list = field(default_factory=list)   # textos para ESTE numero (se envian juntos)
    accion: dict | None = None                     # {"tipo": "registrar"|"entregar", ...}


@dataclass
class Conversacion:
    numero: str
    estado: Estado = Estado.INICIO
    nombre: str | None = None
    lineas: list = field(default_factory=list)     # [{"codigo", "nombre", "cantidad"}]
    menu: list = field(default_factory=list)       # items mostrados en el ultimo catalogo
    pendiente: dict | None = None                  # producto elegido, esperando cantidad
    entrega: str | None = None                     # "envio" | "retiro"
    ubicacion: tuple | None = None
    km: float | None = None
    minutos: int | None = None
    metodo_pago: str | None = None
    total_cotizado: float | None = None
    sinpe: str | None = None
    ultimo_mensaje: datetime | None = None

    def limpiar_pedido(self):
        self.lineas, self.menu, self.pendiente = [], [], None
        self.limpiar_entrega()

    def limpiar_entrega(self):
        self.entrega = self.ubicacion = self.km = self.minutos = None
        self.metodo_pago = self.total_cotizado = None


class GestorConversaciones:
    """Guarda las conversaciones en memoria y aplica el vencimiento por inactividad."""

    def __init__(self, minutos_inactividad=30):
        self.minutos_inactividad = minutos_inactividad
        self.vencimiento = timedelta(minutes=minutos_inactividad)
        self._conversaciones = {}
        self.lock = RLock()          # el webhook y la fila de escrituras usan el mismo lock

    def obtener(self, numero, ahora):
        """Devuelve (conversacion, vencida). Vencida = tenia un pedido a medio hacer y se borro."""
        conv = self._conversaciones.get(numero)
        vencida = False
        if conv and conv.estado not in (Estado.INICIO, Estado.REGISTRANDO) and conv.ultimo_mensaje \
                and ahora - conv.ultimo_mensaje > self.vencimiento:
            if conv.lineas:              # solo se avisa (y se cuenta) si habia productos elegidos
                log.info("PEDIDO_ABANDONADO estado=%s lineas=%s", conv.estado.name, len(conv.lineas))
                vencida = True
            conv = None
        if conv is None:
            conv = Conversacion(numero=numero)
            self._conversaciones[numero] = conv
        return conv, vencida

    def buscar(self, numero):
        return self._conversaciones.get(numero)


# ----------------------------------------------------------------------------
# Textos
# ----------------------------------------------------------------------------
def colones(monto):
    return f"₡{float(monto):,.2f}"


def hora_texto(momento):
    if not isinstance(momento, datetime):
        return str(momento)
    h = momento.hour % 12 or 12
    return (f"{momento.day}-{MESES[momento.month - 1]} {h}:{momento.minute:02d} "
            f"{'a. m.' if momento.hour < 12 else 'p. m.'}")


def _simple(texto):
    t = unicodedata.normalize("NFKD", str(texto or ""))
    return " ".join("".join(c for c in t if not unicodedata.combining(c)).lower().split())


def _opcion(texto):
    t = str(texto or "").strip().rstrip(".").strip()
    return int(t) if t.isdigit() else None


def _saludo(nombre, promocion, vencida, minutos):
    partes = []
    if vencida:
        partes.append(f"Tu pedido anterior quedó sin terminar por más de {minutos} minutos y se canceló. "
                      "¡Empecemos de nuevo!")
    saludo = f"¡Hola, {nombre}! 👋" if nombre else "¡Hola! 👋"
    partes.append(f"{saludo} Bienvenido a la panadería.")
    if promocion:
        partes.append(promocion)
    return "\n\n".join(partes)


def _texto_catalogo(cat):
    lineas = ["🍞 *Menú de hoy* (responde con el número)"]
    if cat.get("descuento_general"):
        pct = round(cat.get("pct_general", 0) * 100)
        lineas.append(f"Hoy es {cat.get('motivo')}: {pct}% de descuento en todos los productos "
                      "(hoy no hay combos).")
    for i, item in enumerate(cat["items"], start=1):
        lineas.append(f"{i}. {item['nombre']} — {colones(item['precio'])}")
        if item.get("detalle"):
            lineas.append(f"    {item['detalle']}")
    lineas.append("0. Cancelar")
    return "\n".join(lineas)


def _texto_algo_mas(conv):
    carrito = "\n".join(f"• {l['cantidad']} {l['nombre']}" for l in conv.lineas)
    return (f"🛒 *Tu pedido hasta ahora*\n{carrito}\n\n¿Qué deseas hacer?\n"
            "1. Agregar otro producto\n2. Quitar un producto\n3. Continuar")


TEXTO_ENTREGA = "🚚 ¿Cómo lo recibes?\n1. Envío a domicilio\n2. Retiro en la panadería"
TEXTO_UBICACION = ("📍 Comparte tu ubicación: toca el clip 📎 → *Ubicación* → *Enviar tu ubicación actual*.\n"
                   "Si prefieres retirar en la panadería, escribe 2.")
TEXTO_PAGO = "💳 ¿Cómo vas a pagar?\n1. Efectivo\n2. SINPE Móvil"


def _texto_resumen(conv, cot):
    partes = ["🧾 *Tu pedido*", "\n".join(f"• {c} {n}" for c, n, _ in cot["lineas"]),
              f"Subtotal: {colones(cot['subtotal_bruto'])}"]
    if cot["descuento_total"]:
        partes.append(f"Descuento: −{colones(cot['descuento_total'])}")
    if conv.entrega == "envio":
        partes.append(f"Envío ({cot['km']} km): {colones(cot['envio'])}")
    else:
        partes.append("Retiro en la panadería (sin costo de envío)")
    partes.append(f"*Total a pagar: {colones(cot['total'])}*")
    if conv.metodo_pago == "Efectivo":
        partes.append("Pago: Efectivo " + ("(se cobra al entregar)" if conv.entrega == "envio"
                                           else "(se paga al retirar)"))
    else:
        partes.append("Pago: SINPE Móvil" + (f" al {conv.sinpe}" if conv.sinpe else ""))
    return ("\n".join(partes) + f"\n\n{AVISO_28}\n\n¿Confirmas tu pedido?\n"
            "1. Sí, confirmar\n2. Cambiar pedido\n3. Cancelar")


def texto_recibo(r):
    partes = [f"🧾 *Recibo — Pedido {r['id_pedido']}*",
              "\n".join(f"• {c} {n}: {colones(s)}" for c, n, s in r["lineas"]),
              f"Subtotal: {colones(r['subtotal_bruto'])}"]
    if r["descuento_total"]:
        partes.append(f"Descuento: −{colones(r['descuento_total'])}")
    partes.append(f"Envío: {colones(r['envio'])}" if r["envio"] else "Retiro en la panadería")
    partes += [f"*Total pagado: {colones(r['total'])}*", f"Forma de pago: {r['metodo_pago']}",
               f"Entregado: {hora_texto(r['entregado_en'])}", "", "¡Gracias por tu compra! 🍞", TEXTO_AYUDA]
    return "\n".join(partes)


# ----------------------------------------------------------------------------
# Logica
# ----------------------------------------------------------------------------
# Acepta signos alrededor: "entregado 120", "Entregado 120:", "ENTREGADO: WA-120.", "entregado wa 120!"
RE_ENTREGADO = re.compile(r"^\W*entregado\W*(?:wa\W*)?(\d+)\W*$", re.IGNORECASE)


def procesar(gestor, msj: MensajeEntrante, servicios, ahora=None) -> Respuesta:
    """Punto de entrada: un mensaje del cliente (o del repartidor) -> Respuesta."""
    ahora = ahora or datetime.now()
    with gestor.lock:
        # Comando del repartidor / administracion (antes que la conversacion normal)
        comando = RE_ENTREGADO.match(msj.texto or "")
        if comando and msj.autorizado:
            id_pedido = f"WA-{comando.group(1)}"
            return Respuesta([f"⏳ Registrando la entrega de {id_pedido}…"],
                             {"tipo": "entregar", "numero": msj.numero, "id_pedido": id_pedido})

        conv, vencida = gestor.obtener(msj.numero, ahora)
        conv.ultimo_mensaje = ahora
        if msj.nombre:
            conv.nombre = msj.nombre
        try:
            return _paso(gestor, conv, msj, servicios, vencida)
        except ErrorParaCliente as error:
            return Respuesta([str(error)])
        except Exception:                                   # nunca dejar al cliente sin respuesta
            log.exception("Error inesperado en la conversacion")
            return Respuesta(["😕 Tuvimos un problema técnico. Por favor intenta de nuevo en unos minutos."])


def _reiniciar(conv):
    conv.limpiar_pedido()
    conv.estado = Estado.INICIO


def _mostrar_catalogo(conv, servicios):
    cat = servicios.catalogo_hoy()
    conv.menu = cat["items"]
    conv.estado = Estado.ELIGIENDO_PRODUCTO
    return _texto_catalogo(cat)


def _cotizar_y_resumir(conv, servicios, aviso=None):
    conv.sinpe = servicios.catalogo_hoy().get("sinpe")
    cot = servicios.cotizar_pedido([{"producto": l["codigo"], "cantidad": l["cantidad"]} for l in conv.lineas],
                                   conv.km, conv.minutos)
    conv.total_cotizado = cot["total"]
    conv.estado = Estado.CONFIRMAR
    texto = _texto_resumen(conv, cot)
    return f"{aviso}\n\n{texto}" if aviso else texto


def _a_pago_retiro(conv):
    conv.entrega, conv.km, conv.minutos, conv.ubicacion = "retiro", None, None, None
    conv.estado = Estado.PAGO
    return Respuesta([TEXTO_PAGO])


def _paso(gestor, conv, msj, servicios, vencida):
    texto, op = msj.texto or "", _opcion(msj.texto)
    e = conv.estado

    if e == Estado.REGISTRANDO:
        return Respuesta(["⏳ Tu pedido se está registrando; te aviso en un momento."])
    if _simple(texto) == "cancelar" and e != Estado.INICIO:
        _reiniciar(conv)
        return Respuesta([TEXTO_CANCELADO])

    if e == Estado.INICIO:                       # saludo + menu de productos en el mismo mensaje
        promo = servicios.catalogo_hoy().get("promocion")
        saludo = _saludo(conv.nombre, promo, vencida, gestor.minutos_inactividad)
        return Respuesta([f"{saludo}\n\n{_mostrar_catalogo(conv, servicios)}"])

    if e == Estado.ELIGIENDO_PRODUCTO:
        if op == 0:
            _reiniciar(conv)
            return Respuesta([TEXTO_CANCELADO])
        if op is None or not 1 <= op <= len(conv.menu):
            return Respuesta([f"Responde con un número del 1 al {len(conv.menu)} (o 0 para cancelar)."])
        conv.pendiente = conv.menu[op - 1]
        conv.estado = Estado.CANTIDAD
        return Respuesta([f"¿Cuántas unidades de *{conv.pendiente['nombre']}*? (1 a {CANTIDAD_MAXIMA})"])

    if e == Estado.CANTIDAD:
        if op is None or not 1 <= op <= CANTIDAD_MAXIMA:
            return Respuesta([f"Escribe una cantidad del 1 al {CANTIDAD_MAXIMA}."])
        item = conv.pendiente
        existente = next((l for l in conv.lineas if l["codigo"] == item["codigo"]), None)
        if existente:
            existente["cantidad"] = min(existente["cantidad"] + op, CANTIDAD_MAXIMA)
        else:
            conv.lineas.append({"codigo": item["codigo"], "nombre": item["nombre"], "cantidad": op})
        conv.pendiente = None
        conv.estado = Estado.ALGO_MAS
        return Respuesta([_texto_algo_mas(conv)])

    if e == Estado.ALGO_MAS:
        if op == 1:
            return Respuesta([_mostrar_catalogo(conv, servicios)])
        if op == 2:
            conv.estado = Estado.QUITANDO
            lista = "\n".join(f"{i}. {l['cantidad']} {l['nombre']}" for i, l in enumerate(conv.lineas, start=1))
            return Respuesta([f"¿Cuál quitas? (responde con el número)\n{lista}\n0. No quitar nada"])
        if op == 3:
            conv.estado = Estado.ENTREGA
            return Respuesta([TEXTO_ENTREGA])
        return Respuesta([_texto_algo_mas(conv)])

    if e == Estado.QUITANDO:
        if op == 0:
            conv.estado = Estado.ALGO_MAS
            return Respuesta([_texto_algo_mas(conv)])
        if op is None or not 1 <= op <= len(conv.lineas):
            return Respuesta([f"Responde con un número del 1 al {len(conv.lineas)} (o 0 para no quitar nada)."])
        quitada = conv.lineas.pop(op - 1)
        if not conv.lineas:
            return Respuesta([f"Quitamos {quitada['nombre']}. Tu pedido quedó vacío.\n\n"
                              + _mostrar_catalogo(conv, servicios)])
        conv.estado = Estado.ALGO_MAS
        return Respuesta([f"Quitamos {quitada['nombre']}.\n\n" + _texto_algo_mas(conv)])

    if e == Estado.ENTREGA:
        if op == 1:
            conv.entrega, conv.estado = "envio", Estado.UBICACION
            return Respuesta([TEXTO_UBICACION])
        if op == 2:
            return _a_pago_retiro(conv)
        return Respuesta([TEXTO_ENTREGA])

    if e == Estado.UBICACION:
        if msj.ubicacion:
            cot = servicios.cotizar_envio(msj.ubicacion)
            if not cot["dentro_zona"]:
                conv.estado = Estado.FUERA_ZONA
                return Respuesta([f"Lo sentimos: estás a {cot['km']} km y entregamos hasta "
                                  f"{cot['radio_maximo_km']:g} km.\n1. Retirar en la panadería\n2. Cancelar pedido"])
            conv.ubicacion, conv.km, conv.minutos = msj.ubicacion, cot["km"], cot["minutos"]
            conv.estado = Estado.PAGO
            return Respuesta([f"Tu dirección está a {cot['km']} km: envío {colones(cot['envio'])}.\n\n{TEXTO_PAGO}"])
        if op == 2:
            return _a_pago_retiro(conv)
        return Respuesta([TEXTO_UBICACION])

    if e == Estado.FUERA_ZONA:
        if op == 1:
            return _a_pago_retiro(conv)
        if op == 2:
            _reiniciar(conv)
            return Respuesta([TEXTO_CANCELADO])
        return Respuesta(["Responde con el número:\n1. Retirar en la panadería\n2. Cancelar pedido"])

    if e == Estado.PAGO:
        if op not in (1, 2):
            return Respuesta([TEXTO_PAGO])
        conv.metodo_pago = "Efectivo" if op == 1 else "SINPE Móvil"
        return Respuesta([_cotizar_y_resumir(conv, servicios)])

    if e == Estado.CONFIRMAR:
        if op == 1:
            conv.estado = Estado.REGISTRANDO
            datos = {
                "lineas": [{"producto": l["codigo"], "cantidad": l["cantidad"]} for l in conv.lineas],
                "id_mensaje": msj.id_mensaje,
                "cliente": conv.nombre,
                "telefono": conv.numero,
                "ubicacion": conv.ubicacion,
                "distancia_km": conv.km,
                "minutos_entrega": conv.minutos,
                "total_cotizado": conv.total_cotizado,
                "metodo_pago": conv.metodo_pago,
            }
            return Respuesta(["⏳ Recibimos tu pedido, lo estamos registrando…"],
                             {"tipo": "registrar", "numero": conv.numero, "datos": datos})
        if op == 2:
            conv.limpiar_entrega()
            conv.estado = Estado.ALGO_MAS
            return Respuesta([_texto_algo_mas(conv)])
        if op == 3:
            _reiniciar(conv)
            return Respuesta([TEXTO_CANCELADO])
        return Respuesta(["Responde con el número:\n1. Sí, confirmar\n2. Cambiar pedido\n3. Cancelar"])

    _reiniciar(conv)
    return Respuesta(["Escribe cualquier cosa para empezar."])


# ----------------------------------------------------------------------------
# Resultados de la fila de escrituras -> mensajes [(numero, texto), ...]
# ----------------------------------------------------------------------------
def tras_registro(gestor, numero, servicios, resultado=None, error=None, minutos_entrega=25):
    with gestor.lock:
        conv = gestor.buscar(numero)
        if error is None and resultado and resultado.get("estado") in ("REGISTRADO", "DUPLICADO"):
            id_pedido = resultado["id_pedido"]
            partes = [f"✅ ¡Listo! Tu pedido *{id_pedido}* quedó registrado."]
            if conv and conv.entrega == "retiro":
                partes.append("Puedes pasar a retirarlo a la panadería; al entregártelo te enviaremos "
                              "tu recibo por aquí.")
            else:
                partes.append(f"Tiempo estimado de entrega: unos {minutos_entrega} minutos.")
                partes.append("Al entregarlo te enviaremos tu recibo por aquí.")
            if conv and conv.metodo_pago == "SINPE Móvil" and resultado.get("total") is not None:
                partes.append(f"Recuerda hacer el SINPE Móvil de {colones(resultado['total'])}"
                              + (f" al {conv.sinpe}" if conv.sinpe else "") + f" e indicar {id_pedido} en el detalle.")
            partes.append(TEXTO_AYUDA)
            if conv:
                _reiniciar(conv)
            return [(numero, "\n".join(partes))]

        if conv is None:
            return [(numero, "😕 No pudimos registrar tu pedido. Por favor intenta de nuevo.")]
        if isinstance(error, TotalCambiado):
            try:
                return [(numero, _cotizar_y_resumir(conv, servicios, aviso="⚠️ El total cambió (por ejemplo, "
                                                     "terminó una promoción). Este es el total actualizado:"))]
            except ErrorParaCliente as otro:
                error = otro
        if isinstance(error, ErrorParaCliente):
            conv.limpiar_entrega()
            conv.estado = Estado.ALGO_MAS
            return [(numero, f"{error}\n\n{_texto_algo_mas(conv)}")]
        log.error("Registro fallido para %s: %r", numero, error)
        conv.estado = Estado.CONFIRMAR
        return [(numero, "😕 No pudimos registrar tu pedido por un problema técnico. "
                         "Escribe 1 para intentar de nuevo o 3 para cancelar.")]


def tras_entrega(numero_repartidor, id_pedido, resultado=None, error=None):
    if error is not None:
        texto = str(error) if isinstance(error, ErrorParaCliente) else f"No se pudo registrar la entrega de {id_pedido}."
        return [(numero_repartidor, f"⚠️ {texto}")]
    if resultado["estado"] == "YA_ENTREGADO":
        return [(numero_repartidor, f"{id_pedido} ya estaba entregado ({hora_texto(resultado['entregado_en'])}). "
                                    "No se envió otro recibo.")]
    if resultado.get("telefono"):
        return [(numero_repartidor, f"✅ {id_pedido} entregado ({hora_texto(resultado['entregado_en'])}). "
                                    "Recibo enviado al cliente."),
                (f"whatsapp:{resultado['telefono']}", texto_recibo(resultado))]
    return [(numero_repartidor, f"✅ {id_pedido} entregado ({hora_texto(resultado['entregado_en'])}). "
                                "El pedido no tiene teléfono: no se envió recibo por WhatsApp.")]
