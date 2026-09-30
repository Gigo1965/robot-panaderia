"""Conector del bot con el Excel maestro (el antes llamado "traductor" del #33).

No tiene reglas de negocio propias: todo lo calcula el robot con las funciones
COMPARTIDAS con el SaaS (Regla #5):
  - catalogo.py            -> productos, combos activos, feriado/dia de pago, SINPE
  - pedidos_digitales.py   -> cotizar_envio, cotizar_pedido, registrar_pedido_digital,
                              marcar_entregado
Lo unico que hace aqui es:
  1) traducir los datos de la conversacion a esas funciones (codigos de la TABLA
     CRUCE MAESTRA, forma de pago, telefono, numero de mensaje);
  2) traducir los errores del robot a mensajes aptos para el cliente.
"""
import logging
import time

from conexion_drive import leer_excel
from microservicios.grupo1_captura_operativa import pedidos_digitales as pd
from microservicios.grupo1_captura_operativa.catalogo import (
    Catalogo, ComboInactivo, PrecioSinCalcular, ProductoNoEncontrado,
)

from . import config
from .conversation import ErrorParaCliente, TotalCambiado

log = logging.getLogger("whatsapp_bot")

ETIQUETA_PROMOCION = "Mensaje de Promoción SaaS / Bot"        # NODO INTEGRAL (columna B -> C)
NO_DISPONIBLE = "😕 Uno de los productos de tu pedido ya no está disponible hoy"


def _numero(valor):
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


def _texto_nodo(wb_valores, etiqueta):
    ws = wb_valores["NODO INTEGRAL"]
    for fila in range(1, ws.max_row + 1):
        texto = ws.cell(fila, 2).value
        if isinstance(texto, str) and texto.strip().lower() == etiqueta.lower():
            valor = ws.cell(fila, 3).value
            return str(valor).strip() if valor not in (None, "") else None
    return None


class ConectorExcel:
    def __init__(self, segundos_cache=None):
        self.segundos_cache = config.SEGUNDOS_CACHE_CATALOGO if segundos_cache is None else segundos_cache
        self._cache = (0.0, None)
        self._autorizados = {self._limpio(n) for n in config.NUMEROS_AUTORIZADOS_TEXTO.split(",") if n.strip()}

    # ------------------------------------------------------------------ lectura
    def catalogo_hoy(self):
        """Lo que se puede vender HOY (se relee del Excel cada pocos minutos)."""
        momento, datos = self._cache
        if datos is not None and time.monotonic() - momento < self.segundos_cache:
            return datos
        wb = leer_excel()
        cat = Catalogo(wb)
        reglas = cat.reglas
        combos = [] if reglas["descuento_general"] else [c for c in cat.combos if c["activo"]]
        items = [{"codigo": c["codigo"], "nombre": c["nombre"], "precio": c["precio"], "es_combo": True,
                  "detalle": c.get("detalle")} for c in combos if _numero(c["precio"])]
        items += [{"codigo": p["codigo"], "nombre": str(p["nombre"]).strip(), "precio": p["precio"],
                   "es_combo": False, "detalle": None} for p in cat.productos if p["codigo"] and _numero(p["precio"])]
        promocion = _texto_nodo(wb, ETIQUETA_PROMOCION)
        if promocion and promocion.startswith("["):          # marcador sin llenar, ej. '[Combo 1'
            promocion = None
        datos = {
            "items": items,
            "descuento_general": reglas["descuento_general"],
            "pct_general": reglas["pct_general"],
            "motivo": "feriado" if reglas["feriado"] else ("día de pago" if reglas["dia_pago"] else None),
            "promocion": promocion,
            "sinpe": reglas.get("sinpe_movil"),
        }
        self._cache = (time.monotonic(), datos)
        return datos

    def cotizar_envio(self, ubicacion):
        try:
            return pd.cotizar_envio(ubicacion)
        except pd.ErrorGoogleMaps as error:
            log.warning("Google Maps fallo: %s", error)
            raise ErrorParaCliente("😕 No pudimos calcular el envío en este momento. Manda tu ubicación de "
                                   "nuevo en un rato, o escribe 2 para retirar en la panadería.") from error

    def cotizar_pedido(self, lineas, km, minutos):
        try:
            return pd.cotizar_pedido(lineas, distancia_km=km, minutos_entrega=minutos)
        except (ProductoNoEncontrado, ComboInactivo, PrecioSinCalcular) as error:
            self._cache = (0.0, None)                              # el catalogo cambio: releer
            raise ErrorParaCliente(f"{NO_DISPONIBLE} ({error}). Quítalo y elige otro.") from error

    # ---------------------------------------------------------------- escritura
    def registrar(self, datos):
        """Lo llama la FILA de escrituras (uno a la vez, en segundo plano)."""
        try:
            return pd.registrar_pedido_digital(
                datos["lineas"], datos["id_mensaje"], cliente=datos["cliente"], metodo_pago=datos["metodo_pago"],
                ubicacion=datos["ubicacion"], distancia_km=datos["distancia_km"],
                minutos_entrega=datos["minutos_entrega"], total_cotizado=datos["total_cotizado"],
                telefono=datos["telefono"],
            )
        except pd.TotalCambio as error:
            self._cache = (0.0, None)
            raise TotalCambiado(str(error)) from error
        except (ProductoNoEncontrado, ComboInactivo, PrecioSinCalcular) as error:
            self._cache = (0.0, None)
            raise ErrorParaCliente(f"{NO_DISPONIBLE} ({error}). Quítalo y elige otro.") from error
        except pd.FueraDeZona as error:
            raise ErrorParaCliente("😕 Tu dirección quedó fuera de la zona de entrega. Puedes retirar en la "
                                   "panadería.") from error

    def entregar(self, id_pedido):
        try:
            return pd.marcar_entregado(id_pedido, origen="WhatsApp")
        except (pd.EntregaNoPermitida, LookupError) as error:
            raise ErrorParaCliente(str(error)) from error

    # ---------------------------------------------------------- autorizaciones
    @staticmethod
    def _limpio(numero):
        try:
            return pd.normalizar_telefono(numero)
        except pd.TelefonoInvalido:
            return None

    def es_autorizado(self, numero):
        limpio = self._limpio(numero)
        return bool(limpio) and limpio in self._autorizados
