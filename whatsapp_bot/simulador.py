"""Simulador del bot en la Terminal: chatea con el bot SIN Twilio, contra la COPIA LOCAL.

Usa el mismo nucleo que el servidor real (conversacion + conector + fila), asi
lo que funciona aqui funciona igual en WhatsApp. OJO: cuando confirmas un pedido
o marcas una entrega, SE ESCRIBE en datos_prueba/maestro_prueba.xlsx
(protocolo: foto antes -> simulacion -> restaurar la copia). Google Maps SI se
consulta de verdad cuando mandas una ubicacion.

Uso (desde la carpeta ROBOT PANADERO):
    py -m whatsapp_bot.simulador

Comandos especiales (empiezan con /):
    /ubicacion canas            manda el pin de un cliente en Canas (dentro de zona)
    /ubicacion liberia          manda el pin de un cliente en Liberia (fuera de zona)
    /ubicacion 10.43,-85.09     manda un pin con esas coordenadas
    /repartidor                 cambia al telefono del repartidor (autorizado solo en el simulador)
    /cliente                    vuelve al telefono del cliente
    /esperar 31                 simula que pasan 31 minutos sin escribir
    /salir                      termina
"""
import os
import sys

# Modo local ANTES de importar el robot (asi nunca toca Drive)
os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import time                                                  # noqa: E402
from datetime import datetime, timedelta                     # noqa: E402

CLIENTE = "whatsapp:+50688887777"          # numeros SUPUESTOS del simulador
REPARTIDOR = "whatsapp:+50670000000"
LUGARES = {"canas": (10.4318, -85.0983), "liberia": (10.6346, -85.4377)}


def main():
    from .conector_excel import ConectorExcel
    from .nucleo import Bot

    class ConectorSimulado(ConectorExcel):
        def es_autorizado(self, numero):            # en el simulador el repartidor es fijo
            return numero == REPARTIDOR or super().es_autorizado(numero)

    def mostrar(numero, texto):
        quien = "cliente" if numero == CLIENTE else ("repartidor" if numero == REPARTIDOR else numero)
        print(f"\n[BOT -> {quien}]\n{texto}\n")

    bot = Bot(enviar=mostrar, servicios=ConectorSimulado(), sincrona=True)
    numero, nombre, desfase, contador = CLIENTE, "Cliente Simulado", timedelta(0), 0
    print(__doc__.split("Comandos especiales")[1])
    print("Escribe como si fueras el cliente. (Copia local: datos_prueba/maestro_prueba.xlsx)\n")

    while True:
        try:
            linea = input(f"TU ({'cliente' if numero == CLIENTE else 'repartidor'}) > ").strip()
        except EOFError:
            break
        ubicacion = None
        if linea.startswith("/"):
            partes = linea[1:].split(maxsplit=1)
            orden, resto = partes[0].lower(), (partes[1] if len(partes) > 1 else "")
            if orden == "salir":
                break
            if orden == "repartidor":
                numero, nombre = REPARTIDOR, "Repartidor"
                print("(ahora escribes desde el telefono del repartidor)")
                continue
            if orden == "cliente":
                numero, nombre = CLIENTE, "Cliente Simulado"
                print("(ahora escribes desde el telefono del cliente)")
                continue
            if orden == "esperar":
                desfase += timedelta(minutes=float(resto or 0))
                print(f"(pasaron {resto} minutos)")
                continue
            if orden == "ubicacion":
                if resto.lower() in LUGARES:
                    ubicacion = LUGARES[resto.lower()]
                else:
                    lat, lng = (float(x) for x in resto.split(","))
                    ubicacion = (lat, lng)
                linea = ""
            else:
                print("(comando desconocido)")
                continue
        contador += 1
        inicio = time.monotonic()
        texto, accion = bot.recibir(numero, linea, nombre=nombre, id_mensaje=f"SIM-{int(time.time())}-{contador}",
                                    ubicacion=ubicacion, ahora=datetime.now() + desfase)
        mostrar(numero, texto)
        if accion:
            print("(escribiendo en la copia local, puede tardar...)")
            bot.programar(accion)
            print(f"(la escritura tardo {time.monotonic() - inicio:.1f} segundos)")


if __name__ == "__main__":
    sys.exit(main())
