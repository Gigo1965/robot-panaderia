"""Prueba de Pedidos digitales SOBRE LA COPIA LOCAL del Excel (no toca Google Drive).

OJO: Google Maps SI se consulta de verdad (3 llamadas; entran en la cuota
gratis). Necesita internet y GOOGLE_MAPS_API_KEY en el .env.

Usa la misma copia que la prueba de Ventas despacho:
    datos_prueba/maestro_prueba.xlsx

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo1_captura_operativa.probar_pedidos_digitales
"""
import os

# Forzar modo local ANTES de importar el resto (asi nunca toca Drive)
os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import logging                                                   # noqa: E402
import time as reloj                                             # noqa: E402

from conexion_drive import leer_excel                            # noqa: E402
from .catalogo import Catalogo, ComboInactivo, ProductoNoEncontrado  # noqa: E402
from .pedidos_digitales import (  # noqa: E402
    FueraDeZona, anular_linea_pedido_digital, cotizar_envio, cotizar_pedido, registrar_pedido_digital,
)

CANAS = (10.4318, -85.0983)       # cliente de prueba dentro de Canas
LIBERIA = (10.6346, -85.4377)     # cliente de prueba fuera de zona

logging.basicConfig(level=logging.WARNING)


def paso(titulo):
    print(f"\n--- {titulo} ---")


def ventas_digitales_hoy():
    """Suma de Flujo caja diario columna C (Ventas digitales) de todo el ano."""
    ws = leer_excel()["Flujo caja diario"]
    return round(sum(ws.cell(f, 3).value or 0 for f in range(5, ws.max_row + 1)
                     if isinstance(ws.cell(f, 3).value, (int, float))), 2)


if __name__ == "__main__":
    id_prueba = f"PRUEBA-{int(reloj.time())}"      # simula el ID unico del mensaje de WhatsApp
    antes = ventas_digitales_hoy()
    print(f"Ventas digitales en Flujo caja ANTES: {antes}")

    catalogo = Catalogo(leer_excel())
    combo = next((c for c in catalogo.combos if c["activo"]), None)

    paso("1a) El cliente comparte su ubicacion (Canas) -> cotizar_envio")
    cot = cotizar_envio(CANAS)
    print(f"  {cot}")

    paso("1b) Antes de confirmar: TOTAL COMPLETO que ve el cliente -> cotizar_pedido")
    lineas = [{"producto_id": "P01", "cantidad": 3}, {"producto_id": "Pa01", "cantidad": 2}]
    if combo:
        lineas.append({"producto_id": combo["codigo"], "cantidad": 1})
    resumen = cotizar_pedido(lineas, distancia_km=cot["km"], minutos_entrega=cot["minutos"])
    for cantidad, nombre, subtotal in resumen["lineas"]:
        print(f"  {cantidad} x {nombre}: {subtotal}")
    print(f"  Envio ({resumen['km']} km, {resumen['minutos']} min): {resumen['envio']}")
    print(f"  TOTAL A PAGAR: {resumen['total']}")

    paso("1c) El cliente dice 'Si' -> registrar_pedido_digital (NO vuelve a llamar a Google)")
    resultado = registrar_pedido_digital(
        lineas, id_mensaje=id_prueba, cliente="Cliente de prueba", ubicacion=CANAS,
        distancia_km=cot["km"], minutos_entrega=cot["minutos"], total_cotizado=resumen["total"],
    )
    for clave, valor in resultado.items():
        print(f"  {clave}: {valor}")
    print("  " + ("OK: total registrado = total mostrado" if resultado["total"] == resumen["total"]
                  else "ERROR: los totales no coinciden"))

    paso("2) El mismo mensaje otra vez (debe decir DUPLICADO)")
    print(" ", registrar_pedido_digital(lineas, id_mensaje=id_prueba))

    paso("3) Producto inexistente (debe rechazarse sin escribir nada)")
    try:
        registrar_pedido_digital([{"producto": "PAN INVENTADO", "cantidad": 1}], id_mensaje=id_prueba + "-b")
    except ProductoNoEncontrado as error:
        print(f"  Rechazado correctamente: {error}")

    paso("4) Combo INACTIVO (debe rechazarse sin escribir nada)")
    inactivo = next((c for c in catalogo.combos if not c["activo"]), None)
    if inactivo:
        try:
            registrar_pedido_digital([{"producto": inactivo["codigo"], "cantidad": 1}], id_mensaje=id_prueba + "-c")
        except ComboInactivo as error:
            print(f"  Rechazado correctamente: {error}")

    paso("4b) Cliente en Liberia (fuera de zona): el bot avisa y no se registra")
    cot_lejos = cotizar_envio(LIBERIA)
    print(f"  cotizar_envio: dentro_zona={cot_lejos['dentro_zona']}, km={cot_lejos['km']}")
    try:
        registrar_pedido_digital(lineas, id_mensaje=id_prueba + "-d", ubicacion=LIBERIA)
    except FueraDeZona as error:
        print(f"  Rechazado correctamente: {error}")

    paso("4c) Retiro en el local (sin ubicacion): envio 0")
    retiro = registrar_pedido_digital([{"producto": "P01", "cantidad": 1}], id_mensaje=id_prueba + "-e")
    print(f"  {retiro['id_pedido']}: envio {retiro['envio']}, total {retiro['total']}")

    paso("5) Anular SOLO la linea de CupCake del pedido")
    print(" ", anular_linea_pedido_digital(resultado["id_pedido"], "Pa01"))

    despues = ventas_digitales_hoy()
    print(f"\nVentas digitales en Flujo caja DESPUES: {despues}  (diferencia {round(despues - antes, 2)})")
    print("Abre datos_prueba/maestro_prueba.xlsx y revisa el final de 'Interface digital',")
    print("'Detalle comanda digital' y 'RegistroEventos'. En 'NODO INTEGRAL' C9/C10 deben")
    print("estar los minutos y km del pedido de Canas (el de Liberia y el retiro no los cambian).")
