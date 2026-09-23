"""
Máquina de estados de la conversación del bot de pedidos.

Este módulo es intencionalmente independiente de Twilio -- no importa
nada de la librería de Twilio ni hace llamadas HTTP. Solo recibe
"eventos" simples (texto, un ID de opción elegida, o una ubicación) y
decide en qué estado sigue la conversación y qué se le debe responder
al cliente. Esto permite probarlo con datos falsos, sin necesidad de
credenciales reales ni de un servidor corriendo.

REGLA DE ARQUITECTURA (ver Reglas_Globales_del_Proyecto, Regla #4):
esta capa NUNCA escribe en Excel directamente. Cuando el pedido queda
CONFIRMADO, solo entrega un diccionario con el pedido completo -- la
función que de verdad escribe en 'Interface digital' y 'Detalle
comanda digital' es una función Python COMPARTIDA con la futura web
del SaaS (todavía pendiente de construir), no algo que viva aquí.
"""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Optional
import time


class Estado(Enum):
    SALUDO = auto()
    ESPERANDO_UBICACION = auto()
    MOSTRANDO_COMBOS = auto()
    MOSTRANDO_MENU_PRODUCTOS = auto()
    ESPERANDO_CANTIDAD = auto()
    PREGUNTANDO_AGREGAR_MAS = auto()
    CONFIRMANDO_PEDIDO = auto()
    PEDIDO_CONFIRMADO = auto()
    CANCELADO = auto()


@dataclass
class LineaPedido:
    producto_id: str
    producto_nombre: str
    cantidad: int


@dataclass
class ConversacionCliente:
    """Estado completo de la conversación de UN cliente (por número de WhatsApp)."""

    numero_whatsapp: str
    estado: Estado = Estado.SALUDO
    ultima_actividad: float = field(default_factory=time.time)

    ubicacion_lat: Optional[float] = None
    ubicacion_lng: Optional[float] = None

    lineas_pedido: list[LineaPedido] = field(default_factory=list)
    producto_id_en_espera_de_cantidad: Optional[str] = None
    producto_nombre_en_espera_de_cantidad: Optional[str] = None

    nombre_cliente: Optional[str] = None

    def tocar(self) -> None:
        self.ultima_actividad = time.time()

    def esta_vencida(self, timeout_segundos: int) -> bool:
        return (time.time() - self.ultima_actividad) > timeout_segundos


class GestorConversaciones:
    """
    Guarda una ConversacionCliente por número de WhatsApp, en memoria.

    NOTA para producción: esto se pierde si el proceso se reinicia. Para
    un volumen pequeño (una panadería) puede ser suficiente para
    empezar, pero si se necesita resistencia a reinicios, esto debería
    respaldarse en una base de datos simple (ej. SQLite o Redis) sin
    cambiar la interfaz pública de esta clase.
    """

    def __init__(self) -> None:
        self._conversaciones: dict[str, ConversacionCliente] = {}

    def obtener_o_crear(self, numero_whatsapp: str) -> ConversacionCliente:
        conv = self._conversaciones.get(numero_whatsapp)
        if conv is None:
            conv = ConversacionCliente(numero_whatsapp=numero_whatsapp)
            self._conversaciones[numero_whatsapp] = conv
        return conv

    def reiniciar(self, numero_whatsapp: str) -> None:
        self._conversaciones.pop(numero_whatsapp, None)

    def limpiar_vencidas(self, timeout_segundos: int) -> int:
        """Elimina conversaciones abandonadas. Devuelve cuántas se limpiaron."""
        vencidas = [
            numero
            for numero, conv in self._conversaciones.items()
            if conv.esta_vencida(timeout_segundos)
        ]
        for numero in vencidas:
            del self._conversaciones[numero]
        return len(vencidas)


# --- Resultado de procesar un evento: qué responder y qué acción tomar ---

@dataclass
class RespuestaBot:
    tipo_mensaje: str  # "texto", "menu_principal", "pedir_ubicacion", "lista_combos",
                        # "lista_productos", "pedir_cantidad", "preguntar_agregar_mas",
                        # "resumen_confirmacion", "pedido_escrito", "cancelado"
    texto: Optional[str] = None
    datos_extra: Optional[dict] = None


def procesar_mensaje_entrante(
    conv: ConversacionCliente,
    texto_libre: Optional[str] = None,
    opcion_elegida_id: Optional[str] = None,
    ubicacion: Optional[tuple[float, float]] = None,
    combos_activos: Optional[list[dict]] = None,
) -> RespuestaBot:
    """
    Punto de entrada único: recibe el evento del cliente y avanza la
    conversación. `app.py` traduce el webhook real de Twilio a estos
    3 tipos de entrada simples (texto_libre, opcion_elegida_id,
    ubicacion) antes de llamar esta función.

    RECORDATORIO DE ARQUITECTURA (16-sep-2026, confirmado por el
    usuario): el cliente NUNCA arma el pedido escribiendo texto libre
    con códigos -- todo es por selección guiada (listas y botones).
    `texto_libre` aquí solo se usa para el saludo inicial y para
    escribir la cantidad numérica cuando se pide explícitamente.
    """
    conv.tocar()

    if conv.estado == Estado.SALUDO:
        # Cualquier mensaje inicial dispara el menú principal.
        conv.estado = Estado.ESPERANDO_UBICACION
        return RespuestaBot(tipo_mensaje="menu_principal_y_pedir_ubicacion")

    if conv.estado == Estado.ESPERANDO_UBICACION:
        if ubicacion is not None:
            conv.ubicacion_lat, conv.ubicacion_lng = ubicacion
            hay_combos = bool(combos_activos)
            conv.estado = Estado.MOSTRANDO_COMBOS if hay_combos else Estado.MOSTRANDO_MENU_PRODUCTOS
            return RespuestaBot(
                tipo_mensaje="lista_combos" if hay_combos else "lista_productos",
                datos_extra={"combos_activos": combos_activos or []},
            )
        return RespuestaBot(
            tipo_mensaje="texto",
            texto="Necesito tu ubicación para calcular el envío. Compártela usando el clip 📎 → Ubicación.",
        )

    if conv.estado == Estado.MOSTRANDO_COMBOS:
        if opcion_elegida_id == "VER_PRODUCTOS_INDIVIDUALES":
            conv.estado = Estado.MOSTRANDO_MENU_PRODUCTOS
            return RespuestaBot(tipo_mensaje="lista_productos")
        if opcion_elegida_id:
            # Se eligió un combo directamente: se agrega como una línea especial.
            conv.lineas_pedido.append(
                LineaPedido(producto_id=opcion_elegida_id, producto_nombre=opcion_elegida_id, cantidad=1)
            )
            conv.estado = Estado.PREGUNTANDO_AGREGAR_MAS
            return RespuestaBot(tipo_mensaje="preguntar_agregar_mas")
        return RespuestaBot(tipo_mensaje="lista_combos")

    if conv.estado == Estado.MOSTRANDO_MENU_PRODUCTOS:
        if opcion_elegida_id:
            conv.producto_id_en_espera_de_cantidad = opcion_elegida_id
            conv.producto_nombre_en_espera_de_cantidad = opcion_elegida_id
            conv.estado = Estado.ESPERANDO_CANTIDAD
            return RespuestaBot(tipo_mensaje="pedir_cantidad", datos_extra={"producto_id": opcion_elegida_id})
        return RespuestaBot(tipo_mensaje="lista_productos")

    if conv.estado == Estado.ESPERANDO_CANTIDAD:
        cantidad = _extraer_cantidad(texto_libre)
        if cantidad is None or cantidad <= 0:
            return RespuestaBot(tipo_mensaje="texto", texto="Por favor indica un número válido de unidades.")
        conv.lineas_pedido.append(
            LineaPedido(
                producto_id=conv.producto_id_en_espera_de_cantidad,
                producto_nombre=conv.producto_nombre_en_espera_de_cantidad,
                cantidad=cantidad,
            )
        )
        conv.producto_id_en_espera_de_cantidad = None
        conv.producto_nombre_en_espera_de_cantidad = None
        conv.estado = Estado.PREGUNTANDO_AGREGAR_MAS
        return RespuestaBot(tipo_mensaje="preguntar_agregar_mas")

    if conv.estado == Estado.PREGUNTANDO_AGREGAR_MAS:
        if opcion_elegida_id == "AGREGAR_OTRO":
            conv.estado = Estado.MOSTRANDO_MENU_PRODUCTOS
            return RespuestaBot(tipo_mensaje="lista_productos")
        if opcion_elegida_id == "TERMINAR_PEDIDO":
            conv.estado = Estado.CONFIRMANDO_PEDIDO
            return RespuestaBot(
                tipo_mensaje="resumen_confirmacion",
                datos_extra={"lineas": [l.__dict__ for l in conv.lineas_pedido]},
            )
        return RespuestaBot(tipo_mensaje="preguntar_agregar_mas")

    if conv.estado == Estado.CONFIRMANDO_PEDIDO:
        if opcion_elegida_id == "CONFIRMAR":
            conv.estado = Estado.PEDIDO_CONFIRMADO
            pedido_final = {
                "numero_whatsapp": conv.numero_whatsapp,
                "ubicacion_lat": conv.ubicacion_lat,
                "ubicacion_lng": conv.ubicacion_lng,
                "lineas": [l.__dict__ for l in conv.lineas_pedido],
            }
            return RespuestaBot(tipo_mensaje="pedido_listo_para_escribir", datos_extra=pedido_final)
        if opcion_elegida_id == "CANCELAR":
            conv.estado = Estado.CANCELADO
            return RespuestaBot(tipo_mensaje="cancelado")
        return RespuestaBot(
            tipo_mensaje="resumen_confirmacion",
            datos_extra={"lineas": [l.__dict__ for l in conv.lineas_pedido]},
        )

    # Estados terminales (PEDIDO_CONFIRMADO, CANCELADO): cualquier mensaje nuevo reinicia.
    return RespuestaBot(tipo_mensaje="conversacion_terminada")


def _extraer_cantidad(texto: Optional[str]) -> Optional[int]:
    if not texto:
        return None
    texto = texto.strip()
    if texto.isdigit():
        return int(texto)
    return None
