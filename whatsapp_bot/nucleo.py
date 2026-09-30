"""Nucleo del bot: une la conversacion, el conector con el Excel y la fila de escrituras.

Lo usan app.py (Twilio real) y simulador.py (Terminal), asi los dos se comportan
EXACTAMENTE igual; solo cambia como se envian los mensajes.
"""
import logging
from datetime import datetime

from . import config
from .conversation import GestorConversaciones, MensajeEntrante, procesar, tras_entrega, tras_registro
from .fila_escrituras import FilaEscrituras

log = logging.getLogger("whatsapp_bot")


class Bot:
    def __init__(self, enviar, servicios=None, sincrona=False):
        """enviar(numero, texto): como mandar un mensaje FUERA de la respuesta inmediata
        (confirmacion con el numero de pedido, recibo, aviso al repartidor)."""
        if servicios is None:
            from .conector_excel import ConectorExcel
            servicios = ConectorExcel()
        self.servicios = servicios
        self.enviar = enviar
        self.gestor = GestorConversaciones(config.MINUTOS_INACTIVIDAD)
        self.fila = FilaEscrituras(sincrona=sincrona)

    def recibir(self, numero, texto="", nombre=None, id_mensaje=None, ubicacion=None, ahora=None):
        """Devuelve (respuesta_inmediata, accion). Quien llama envia la respuesta y
        DESPUES llama programar(accion)."""
        msj = MensajeEntrante(numero=numero, texto=texto or "", nombre=nombre, id_mensaje=id_mensaje,
                              ubicacion=ubicacion, autorizado=self.servicios.es_autorizado(numero))
        respuesta = procesar(self.gestor, msj, self.servicios, ahora or datetime.now())
        return "\n\n".join(respuesta.mensajes), respuesta.accion

    def programar(self, accion):
        if not accion:
            return
        if accion["tipo"] == "registrar":
            numero, datos = accion["numero"], accion["datos"]

            def al_terminar(resultado, error):
                for destino, texto in tras_registro(self.gestor, numero, self.servicios, resultado, error,
                                                    config.MINUTOS_ENTREGA_ESTIMADOS):
                    self.enviar(destino, texto)

            self.fila.encolar(lambda: self.servicios.registrar(datos), al_terminar)

        elif accion["tipo"] == "entregar":
            numero, id_pedido = accion["numero"], accion["id_pedido"]

            def al_terminar(resultado, error):
                for destino, texto in tras_entrega(numero, id_pedido, resultado, error):
                    self.enviar(destino, texto)

            self.fila.encolar(lambda: self.servicios.entregar(id_pedido), al_terminar)
