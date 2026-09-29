"""Prueba de Hornadas SOBRE LA COPIA LOCAL del Excel (no toca Google Drive).

Antes, la copia debe tener aplicadas las migraciones 001 y 002:
    py -m migraciones.m001_grupo2 --local
    py -m migraciones.m002_consumo_mp --local

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo2_inventario.probar_hornadas

|| V15 (28-sep-2026, Decisiones Pendientes #29): etiquetas con el nombre unico de la
TABLA CRUCE MAESTRA (Regla #26: 'Baguette', 'CupCake'); levadura y leche vigiladas con
su nombre actual en Control de Inventario (antes 'LEVADURA TEMPO SALADA 500 GR' y
'Leche Suli...', que ya no existen y no se mostraban); numeros esperados de levadura
con el paquete de 400 g ('Consumo de Materia Prima' F9/F15).
|| V15 (28-sep-2026, Regla #34, opcion C): paso 0 = una hornada SIN materia prima
suficiente se rechaza sin escribir (incluye las bolsas); paso 0b = compra de prueba
(ingreso_inventario, costo congelado) para que las hornadas de los pasos 1 y 3 alcancen.
Al final debe decir "RESULTADO: TODO OK (0 lineas ERROR)" (los pasos 1 a 7 se siguen
comparando a ojo, como antes).
"""
import os

# Forzar modo local ANTES de importar el resto (asi nunca toca Drive)
os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import logging                                                  # noqa: E402
import time as reloj                                            # noqa: E402

from conexion_drive import leer_excel                           # noqa: E402
from conexion_drive.tiempo import ahora_local                   # noqa: E402
from microservicios.grupo1_captura_operativa.catalogo import ComboInactivo  # noqa: E402
from microservicios.grupo1_captura_operativa.ventas_despacho import registrar_venta_despacho  # noqa: E402
from microservicios.grupo1_captura_operativa.ingreso_inventario import registrar_ingresos  # noqa: E402
from .hornadas import HornadaInvalida, MateriaPrimaInsuficiente, anular_hornada, registrar_hornada  # noqa: E402

logging.basicConfig(level=logging.WARNING)
MESES = "BCDEFGHIJKLM"   # columnas de enero a diciembre en 'flujo caja anual'


# V15: nombres actuales de Control de Inventario (filas 13, 17 y 7)
MATERIAS = ("HARINA FLORES 25 KG FUERTE", "LEVADURA FRESCA GOLONDRINA 400 GR CIP", "Leche Semidescremada Dos Pinos -1 L")


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
    return existencia, {"cantidad Baguette (mes)": anual[f"{col}9"].value,     # V15: Regla #26
                        "cantidad CupCake (mes)": anual[f"{col}10"].value}


def paso(titulo):
    print(f"\n--- {titulo} ---")


ERRORES = []


def revisar(condicion, texto):
    """Imprime OK o ERROR y cuenta los errores para el resumen final (V15)."""
    print(f"  {'OK   ' if condicion else 'ERROR'} {texto}")
    if not condicion:
        ERRORES.append(texto)


def filas_historico():
    return leer_excel()["Historico Diario"].max_row


# V15: compra de prueba para que alcancen 120 Baguette + 40 CupCake (Regla #34).
# Con la copia igual a Drive v97 faltan harina, mejorante, levadura, bolsas (0) y leche.
COMPRA_PRUEBA = [{"producto": "Ha01", "cantidad": 2}, {"producto": "Me01", "cantidad": 1},
                 {"producto": "Le01", "cantidad": 2}, {"producto": "Bo01", "cantidad": 2},
                 {"producto": "L01", "cantidad": 1}]


if __name__ == "__main__":
    marca = int(reloj.time()) % 100000
    print("Materia prima al INICIO:", existencias_mp())

    paso("0) V15 Regla #34: hornada SIN materia prima suficiente -> se rechaza sin escribir")
    antes_hist = filas_historico()
    try:
        registrar_hornada("P01", 120)
        revisar(False, "120 Baguette sin materia prima suficiente debio rechazarse")
    except MateriaPrimaInsuficiente as error:
        print(f"  Rechazado: {error}")
        materias = [f["materia"] for f in error.faltantes]
        revisar(any("BOLSAS" in m for m in materias), f"los empaques cuentan (bolsas en la lista): {len(materias)} faltantes")
        revisar(isinstance(error, HornadaInvalida), "MateriaPrimaInsuficiente es una HornadaInvalida (el SaaS la maneja igual)")
    revisar(filas_historico() == antes_hist, f"no se escribio nada en Historico Diario (max_row sigue en {antes_hist})")

    paso("0b) Compra de prueba de materia prima (primero se compra, despues se hornea)")
    compra = registrar_ingresos(COMPRA_PRUEBA, factura=f"F-PRUEBA-HORNADAS-{marca}", proveedor="Proveedor A")
    revisar(compra.get("estado") == "REGISTRADO" and len(compra.get("registradas", [])) == len(COMPRA_PRUEBA),
            f"compra registrada: {len(compra.get('registradas', []))} lineas")

    print("\nANTES (despues de la compra):", *foto())
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
    print("  Harina: 120 x 230 g + 40 x 25 g = 28.6 kg = 1.144 sacos menos. Levadura: 414 g + 60 g = 474 g = 1.185 paquetes de 400 g.")
    print("  Leche: 40 CupCake x 15 ml = 0.6 litros menos.")

    paso("5) Anular la hornada de Baguette (G2 = 0 + correccion negativa)")
    print(" ", anular_hornada("P01"))

    paso("6) Anularla otra vez (debe decir YA_ESTABA_ANULADA)")
    print(" ", anular_hornada("P01"))

    print("\nDESPUES de anular:", *foto())
    print("  Baguette debe volver a como estaba ANTES; CupCake se queda con +40.")
    print("Materia prima DESPUES de anular la Baguette:", existencias_mp())
    print("  Solo queda rebajado lo de los 40 CupCake (harina 0.04 sacos, levadura 0.15 paquetes, leche 0.6).")

    paso("7) Combos descuentan inventario: vender 1 Combo 2 en despacho (2 CupCake + 1 Coca Cola)")
    print("  ANTES   (existencia, salidas en combos):", existencias_combo())
    try:
        registrar_venta_despacho("co02", 1)
        print("  DESPUES (existencia, salidas en combos):", existencias_combo())
        print("  CupCake debe bajar 2 y la Coca Cola 1.")
    except ComboInactivo as error:
        print(f"  Hoy el Combo 2 no esta activo, no se puede probar este paso: {error}")

    if ERRORES:
        print(f"\nRESULTADO: HAY {len(ERRORES)} lineas ERROR:")
        for texto in ERRORES:
            print("  -", texto)
    else:
        print("\nRESULTADO: TODO OK (0 lineas ERROR)")
