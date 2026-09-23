"""Prueba de Hornadas SOBRE LA COPIA LOCAL del Excel (no toca Google Drive).

Antes, la copia debe tener aplicadas las migraciones 001 y 002:
    py -m migraciones.m001_grupo2 --local
    py -m migraciones.m002_consumo_mp --local

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo2_inventario.probar_hornadas
"""
import os

# Forzar modo local ANTES de importar el resto (asi nunca toca Drive)
os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import logging                                                  # noqa: E402

from conexion_drive import leer_excel                           # noqa: E402
from conexion_drive.tiempo import ahora_local                   # noqa: E402
from microservicios.grupo1_captura_operativa.catalogo import ComboInactivo  # noqa: E402
from microservicios.grupo1_captura_operativa.ventas_despacho import registrar_venta_despacho  # noqa: E402
from .hornadas import HornadaInvalida, anular_hornada, registrar_hornada  # noqa: E402

logging.basicConfig(level=logging.WARNING)
MESES = "BCDEFGHIJKLM"   # columnas de enero a diciembre en 'flujo caja anual'


MATERIAS = ("HARINA FLORES 25 KG FUERTE", "LEVADURA TEMPO SALADA 500 GR", "Leche Suli Semidescremada 1.5% grasa  - 1 L")


def existencias_mp():
    ctrl = leer_excel()["Control de Inventario"]
    return {ctrl.cell(f, 1).value[:22]: round(ctrl.cell(f, 6).value, 4) for f in range(5, ctrl.max_row + 1)
            if ctrl.cell(f, 1).value in MATERIAS}


def existencias_combo():
    ctrl = leer_excel()["Control de Inventario"]
    return {ctrl.cell(f, 1).value: (ctrl.cell(f, 6).value, ctrl.cell(f, 9).value)
            for f in range(5, ctrl.max_row + 1)
            if ctrl.cell(f, 1).value in ("CupCake", "Gaseosa Coca Cola Regular - 600 ml")}


def foto():
    wb = leer_excel()
    ctrl = wb["Control de Inventario"]
    existencia = {ctrl.cell(f, 1).value: ctrl.cell(f, 6).value for f in range(5, ctrl.max_row + 1)
                  if ctrl.cell(f, 1).value in ("Baguette", "CupCake")}
    col = MESES[ahora_local().month - 1]
    anual = wb["flujo caja anual"]
    return existencia, {"cantidad pan baguette (mes)": anual[f"{col}9"].value,
                        "cantidad cupcakes (mes)": anual[f"{col}10"].value}


def paso(titulo):
    print(f"\n--- {titulo} ---")


if __name__ == "__main__":
    print("ANTES:", *foto())
    print("Materia prima ANTES:", existencias_mp())

    paso("1) Hornada nueva: 120 Baguette")
    momento = ahora_local()
    print(" ", registrar_hornada("P01", 120, momento))

    paso("2) La misma hornada otra vez (debe decir DUPLICADO)")
    print(" ", registrar_hornada("P01", 120, momento))

    paso("3) Hornada nueva: 40 CupCake")
    print(" ", registrar_hornada("Pa01", 40))

    paso("4) Producto que no es de produccion propia (debe rechazarse)")
    try:
        registrar_hornada("L01", 10)
    except HornadaInvalida as error:
        print(f"  Rechazado correctamente: {error}")

    print("\nDESPUES de registrar:", *foto())
    print("  Baguette debe subir 120 y CupCake 40, en inventario y en el mes de 'flujo caja anual'.")
    print("Materia prima DESPUES de hornear:", existencias_mp())
    print("  Harina: 120 x 230 g + 40 x 25 g = 28.6 kg = 1.144 sacos menos. Levadura: 414 g + 60 g = 0.948 paquetes.")
    print("  Leche: 40 CupCake x 15 ml = 0.6 litros menos.")

    paso("5) Anular la hornada de Baguette (G2 = 0 + correccion negativa)")
    print(" ", anular_hornada("P01"))

    paso("6) Anularla otra vez (debe decir YA_ESTABA_ANULADA)")
    print(" ", anular_hornada("P01"))

    print("\nDESPUES de anular:", *foto())
    print("  Baguette debe volver a como estaba ANTES; CupCake se queda con +40.")
    print("Materia prima DESPUES de anular la Baguette:", existencias_mp())
    print("  Solo queda rebajado lo de los 40 CupCake (harina 0.04 sacos, levadura 0.12, leche 0.6).")

    paso("7) Combos descuentan inventario: vender 1 Combo 2 en despacho (2 CupCake + 1 Coca Cola)")
    print("  ANTES   (existencia, salidas en combos):", existencias_combo())
    try:
        registrar_venta_despacho("co02", 1)
        print("  DESPUES (existencia, salidas en combos):", existencias_combo())
        print("  CupCake debe bajar 2 y la Coca Cola 1.")
    except ComboInactivo as error:
        print(f"  Hoy el Combo 2 no esta activo, no se puede probar este paso: {error}")
