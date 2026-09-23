"""Prueba de Ventas despacho SOBRE UNA COPIA LOCAL del Excel (no toca Google Drive).

Antes de correrla:
  1. Crea la carpeta  datos_prueba  dentro de ROBOT_PANADERIA
  2. Copia ahi el Excel maestro y renombralo:  maestro_prueba.xlsx

Uso (desde la carpeta ROBOT_PANADERIA):
    python -m microservicios.grupo1_captura_operativa.probar_ventas_despacho
"""
import os

# Forzar modo local ANTES de importar el resto (asi nunca toca Drive)
os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import logging                                                    # noqa: E402

from conexion_drive import leer_excel                             # noqa: E402
from conexion_drive.tiempo import ahora_local                     # noqa: E402
from .catalogo import Catalogo, ComboInactivo, ProductoNoEncontrado  # noqa: E402
from .ventas_despacho import anular_venta_despacho, registrar_venta_despacho  # noqa: E402

logging.basicConfig(level=logging.WARNING)


def total_ventas():
    ws = leer_excel()["Ventas despacho"]
    return ws["N5"].value, ws["N6"].value


def paso(titulo):
    print(f"\n--- {titulo} ---")


if __name__ == "__main__":
    momento = ahora_local()
    antes = total_ventas()
    print(f"Total ventas / anulado ANTES: {antes}")

    paso("Productos y combos vendibles hoy")
    for codigo, nombre, precio in Catalogo(leer_excel()).lista_para_mostrar():
        print(f"  {codigo:<6} {nombre[:45]:<45} {precio:>10.2f}")

    paso("1) Venta nueva: 3 x P01 (Baguette)")
    print(registrar_venta_despacho("P01", 3, momento))

    paso("2) La misma venta otra vez (debe decir DUPLICADO)")
    print(registrar_venta_despacho("P01", 3, momento))

    paso("3) Producto que no existe (debe rechazarse)")
    try:
        registrar_venta_despacho("MEJORANTE STAR PAN PLUS 5 KILOS", 1, momento)
    except ProductoNoEncontrado as error:
        print(f"  Rechazado correctamente: {error}")

    paso("4) Combo: el primero ACTIVO o el aviso si esta inactivo")
    catalogo = Catalogo(leer_excel())
    for combo in catalogo.combos:
        try:
            print(registrar_venta_despacho(combo["codigo"], 1, momento))
            break
        except ComboInactivo as error:
            print(f"  {error}")

    paso("5) Anular la venta del paso 1")
    print(anular_venta_despacho("P01", 3, momento))

    print(f"\nTotal ventas / anulado DESPUES: {total_ventas()}")
    print("Abre datos_prueba/maestro_prueba.xlsx y revisa el final de 'Ventas despacho' y 'RegistroEventos'.")
