"""Prueba de Ingreso de Inventario SOBRE LA COPIA LOCAL del Excel (no toca Google Drive).

Usa la misma copia que las otras pruebas:  datos_prueba/maestro_prueba.xlsx

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo1_captura_operativa.probar_ingreso_inventario
"""
import os

# Forzar modo local ANTES de importar el resto (asi nunca toca Drive)
os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import logging                                                   # noqa: E402
import time as reloj                                             # noqa: E402

from conexion_drive import leer_excel                            # noqa: E402
from .ingreso_inventario import (  # noqa: E402
    ProductoNoReconocido, buscar_opciones, leer_productos_inventario, registrar_ingreso, registrar_ingresos,
)

logging.basicConfig(level=logging.WARNING)
VIGILADOS = ("Baguette", "Leche Suli Semidescremada 1.5% grasa  - 1 L")


def existencias():
    ws = leer_excel()["Control de Inventario"]
    return {ws.cell(f, 1).value: ws.cell(f, 6).value for f in range(5, ws.max_row + 1)
            if ws.cell(f, 1).value in VIGILADOS}


def paso(titulo):
    print(f"\n--- {titulo} ---")


if __name__ == "__main__":
    marca = int(reloj.time()) % 100000
    antes = existencias()
    print("Existencia ANTES (Control de Inventario):", antes)

    paso("0) Autocompletado del SaaS: lo que ve la persona mientras escribe")
    productos = leer_productos_inventario(leer_excel())
    for escrito in ("har", "lech", "bag", "mejorante 5", "cafe"):
        opciones = ", ".join(f"{nombre[:32]} ({codigo})" for codigo, nombre in buscar_opciones(productos, escrito, 3))
        print(f"  escribe '{escrito}'  ->  {opciones}")

    paso("1) Factura de compra: 10 Harina (materia prima) + 12 Leche (reventa)")
    print(" ", registrar_ingresos(
        [{"producto": "Ha01", "cantidad": 10}, {"producto": "L01", "cantidad": 12}],
        factura=f"F-PRUEBA-{marca}", proveedor="Proveedor A"))

    paso("2) Lote de produccion propia: 500 Baguette")
    print(" ", registrar_ingreso("P01", 500, factura=f"L-PRUEBA-{marca}"))

    paso("3) La misma factura otra vez (debe decir DUPLICADO)")
    print(" ", registrar_ingresos(
        [{"producto": "Ha01", "cantidad": 10}, {"producto": "L01", "cantidad": 12}],
        factura=f"F-PRUEBA-{marca}", proveedor="Proveedor A"))

    for numero, escrito in ((4, "MEJORANTE STAR PAN PLUS 5 KILOS"), (5, "CAFE NARANJO PURO 250 GR X 10 UNIDADES")):
        paso(f"{numero}) Nombre mal escrito (caso real): debe rechazar y sugerir")
        try:
            registrar_ingreso(escrito, 2, factura=f"F-PRUEBA-{marca}-{numero}", proveedor="Proveedor B")
        except ProductoNoReconocido as error:
            print(f"  Rechazado correctamente: {error}")

    despues = existencias()
    print("\nExistencia DESPUES (Control de Inventario):", despues)
    print("  Baguette debe subir 500 y la Leche 12.")
    print("Abre datos_prueba/maestro_prueba.xlsx y revisa el final de 'Ingreso de Inventario'.")
