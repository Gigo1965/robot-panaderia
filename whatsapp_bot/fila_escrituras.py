"""Fila de escrituras del bot: los cambios al Excel se hacen UNO A LA VEZ, en segundo plano.

Por que: escribir en el Excel (bajar de Drive, escribir, recalcular con LibreOffice,
subir) tarda de 15 segundos a mas de un minuto, y Twilio solo espera unos 15
segundos la respuesta del webhook. Por eso el bot contesta al instante ("lo estamos
registrando...") y deja el trabajo aqui. Un solo trabajador procesa la fila en
orden, asi dos pedidos simultaneos del mismo servidor nunca chocan entre si.
(Los choques con OTROS procesos, ej. el SaaS, los resuelve editar_excel con su
revision de version y reintentos.)
"""
import logging
import queue
import threading

log = logging.getLogger("whatsapp_bot")


class FilaEscrituras:
    def __init__(self, sincrona=False):
        """sincrona=True: ejecuta cada trabajo en el momento (lo usa el simulador)."""
        self.sincrona = sincrona
        self._fila = queue.Queue()
        if not sincrona:
            threading.Thread(target=self._trabajar, name="fila-escrituras", daemon=True).start()

    def encolar(self, trabajo, al_terminar):
        """trabajo() hace la escritura; al_terminar(resultado, error) recibe el resultado."""
        if self.sincrona:
            self._ejecutar(trabajo, al_terminar)
        else:
            self._fila.put((trabajo, al_terminar))
            log.info("Trabajo en la fila (pendientes: %s)", self._fila.qsize())

    def pendientes(self):
        return self._fila.qsize()

    def _trabajar(self):
        while True:
            trabajo, al_terminar = self._fila.get()
            try:
                self._ejecutar(trabajo, al_terminar)
            finally:
                self._fila.task_done()

    @staticmethod
    def _ejecutar(trabajo, al_terminar):
        resultado, error = None, None
        try:
            resultado = trabajo()
        except Exception as e:          # el error viaja a al_terminar, que avisa al cliente
            error = e
            log.warning("La escritura fallo: %r", e)
        try:
            al_terminar(resultado, error)
        except Exception:
            log.exception("Fallo al avisar el resultado de una escritura")
