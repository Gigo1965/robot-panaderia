"""Compra de prueba para la demostracion (opcion B, decision del usuario 28-sep-2026).

Por que existe (Regla #34 + Regla #27):
  Con el Excel de Drive v97 no se puede hornear: bolsas kraft = 0, harina 0.77
  sacos, mejorante 0.1. Como "no hay hornada sin materia prima", primero se
  registra una compra y despues se hornea. Es un estudio academico: la compra es
  un dato supuesto que abastece una "semana de demostracion" (5 hornadas de 100
  Baguette y 5 de 100 CupCake).

Como lo hace:
  NO es una migracion. Usa la funcion REAL del robot, registrar_ingresos()
  (ingreso_inventario.py), igual que lo haria el SaaS. Por eso:
    - el costo de cada compra sale de la TABLA CRUCE MAESTRA / MP y queda
      congelado (Regla #33);
    - si alguna linea falla, no se escribe ninguna (factura completa o nada);
    - si se corre dos veces, la 2a vez responde DUPLICADO y no escribe nada.
  Las rebajas de inventario las sigue haciendo el Excel con sus formulas.

Facturas (fecha fija 2026-09-28, igual en local y en Drive):
  F005  Solera    : Ha01 5, Me01 3, Le01 6, Bo01 5, Ma03 1, Az01 1, Bu01 2
  F006  MaxiPali  : L01 8

Uso (carpeta ROBOT PANADERO, Terminal):
  py compra_demo_semana.py --local   -> copia datos_prueba/maestro_prueba.xlsx
  py compra_demo_semana.py           -> Excel REAL en Google Drive (pide escribir SI)
"""
import os
import sys

LOCAL = "--local" in sys.argv
if LOCAL:
    # Igual que las migraciones: se fija ANTES de importar conexion_drive.
    os.environ["ORIGEN_EXCEL"] = "local"

from conexion_drive import leer_excel  # noqa: E402
from microservicios.grupo1_captura_operativa.ingreso_inventario import registrar_ingresos  # noqa: E402

FECHA = "2026-09-28"

FACTURAS = [
    ("F005", "Solera", [
        {"producto": "Ha01", "cantidad": 5},
        {"producto": "Me01", "cantidad": 3},
        {"producto": "Le01", "cantidad": 6},
        {"producto": "Bo01", "cantidad": 5},
        {"producto": "Ma03", "cantidad": 1},
        {"producto": "Az01", "cantidad": 1},
        {"producto": "Bu01", "cantidad": 2},
    ]),
    ("F006", "MaxiPali", [
        {"producto": "L01", "cantidad": 8},
    ]),
]

VIGILADOS = [
    "HARINA FLORES 25 KG FUERTE",
    "MEJORANTE PANNIVITA PRO-PAN 1 KG HOMO-PLUS",
    "LEVADURA FRESCA GOLONDRINA 400 GR CIP",
    "BOLSAS PAPEL KRAFT 2 BAGUETTE 100 UNDS BIODEGRADABLE",
    "MARGARINA TUMBA NUMAR 7.5 KILOS 60 BARRITAS X125GR",
    "AZUCAR DOÑA MARIA 2KG (24 KG)",
    "Leche Semidescremada Dos Pinos -1 L",
    "BURBUJA FOAM #6 125 UNID POSTRE BL",
    "SAL SARU COMUN ECONOMICA 25 KG SACO BRIN",
    "MANTECA MAX PAN BULK 25 KGS",
]


def foto():
    """SOLO LEE: existencias (Control de Inventario F) y el dia FECHA de 'Flujo caja diario'."""
    wb = leer_excel()
    existencias = {}
    for fila in wb["Control de Inventario"].iter_rows(min_row=5, values_only=True):
        if fila[0]:
            existencias[str(fila[0]).strip()] = fila[5]
    dia = None
    for fila in wb["Flujo caja diario"].iter_rows(min_row=5, values_only=True):
        if hasattr(fila[0], "date") and fila[0].date().isoformat() == FECHA:
            dia = {"E materia prima": fila[4], "F compras externas": fila[5], "H egresos": fila[7],
                   "I flujo del dia": fila[8], "J acumulado": fila[9]}
            break
    return existencias, dia


def _num(valor):
    return f"{valor:,.4f}" if isinstance(valor, (int, float)) else repr(valor)


def main():
    destino = "COPIA LOCAL (datos_prueba/maestro_prueba.xlsx)" if LOCAL else "GOOGLE DRIVE (real)"
    print(f"Destino: {destino}")
    print(f"Fecha de las compras: {FECHA}\n")

    if not LOCAL:
        respuesta = input("Esto escribe en el Excel REAL de Drive. Escriba SI para seguir: ").strip().upper()
        if respuesta != "SI":
            print("Cancelado. No se toco nada.")
            return

    antes, dia_antes = foto()

    for factura, proveedor, lineas in FACTURAS:
        resultado = registrar_ingresos(lineas, factura, proveedor=proveedor, fecha=FECHA)
        print(f"Factura {factura} ({proveedor}): {resultado.get('estado')} - "
              f"registradas {len(resultado.get('registradas', []))}, "
              f"duplicadas {len(resultado.get('duplicadas', []))}")
        for r in resultado.get("registradas", []):
            print(f"   fila {r['fila']}: {r['cantidad']} x {r['producto']} - costo congelado {r['costo_unitario']:,.2f}")

    despues, dia_despues = foto()

    print("\nEXISTENCIAS (Control de Inventario, columna F, en paquetes): antes -> despues")
    for nombre in VIGILADOS:
        print(f"   {nombre}: {_num(antes.get(nombre))} -> {_num(despues.get(nombre))}")

    print(f"\nFLUJO CAJA DIARIO del {FECHA}: antes -> despues")
    if dia_antes is None and dia_despues is None:
        print("   (la fila de ese dia no existe todavia)")
    else:
        for clave in ("E materia prima", "F compras externas", "H egresos", "I flujo del dia", "J acumulado"):
            a = (dia_antes or {}).get(clave)
            d = (dia_despues or {}).get(clave)
            print(f"   {clave}: {_num(a)} -> {_num(d)}")

    print("\nLISTO.")


if __name__ == "__main__":
    main()
