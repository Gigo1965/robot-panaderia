"""Prueba de Pedidos digitales SOBRE LA COPIA LOCAL del Excel (no toca Google Drive).

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
from .pedidos_digitales import anular_linea_pedido_digital, registrar_pedido_digital  # noqa: E402

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

    paso("1) Pedido nuevo: 3 Baguette + 2 CupCake" + (f" + 1 {combo['nombre']}" if combo else ""))
    lineas = [{"producto_id": "P01", "cantidad": 3}, {"producto_id": "Pa01", "cantidad": 2}]
    if combo:
        lineas.append({"producto_id": combo["codigo"], "cantidad": 1})
    resultado = registrar_pedido_digital(
        lineas, id_mensaje=id_prueba, cliente="Cliente de prueba",
        ubicacion=(10.4318, -85.0983), metodo_pago=None,
    )
    for clave, valor in resultado.items():
        print(f"  {clave}: {valor}")

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

    paso("5) Anular SOLO la linea de CupCake del pedido")
    print(" ", anular_linea_pedido_digital(resultado["id_pedido"], "Pa01"))

    despues = ventas_digitales_hoy()
    print(f"\nVentas digitales en Flujo caja DESPUES: {despues}  (diferencia {round(despues - antes, 2)})")
    print("Abre datos_prueba/maestro_prueba.xlsx y revisa el final de 'Interface digital',")
    print("'Detalle comanda digital' y 'RegistroEventos'.")
