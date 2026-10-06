"""Prueba completa de costeo.py sobre una COPIA TEMPORAL del Excel de prueba.

No toca datos_prueba/maestro_prueba.xlsx ni Google Drive: copia el Excel de prueba a
datos_prueba/_prueba_costeo.xlsx (se reemplaza en cada corrida) y ahi ejecuta 14 casos:
7 cambios validos y 7 que DEBEN ser rechazados. Al final compara TODAS las formulas del libro
contra el original (deben quedar identicas) y revisa los resultados calculados.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo4_costeo.probar_costeo
Tarda unos minutos: cada cambio valido se recalcula con LibreOffice, igual que en produccion.

|| 5-oct-2026 (SaaS Grupo 4, #78): caso 15 = un cambio con origen "SaaS (administracion)" (son 15
casos: 8 validos y 7 rechazados). Se revisa que los 7 eventos de antes sigan diciendo "Robot Python",
que el nuevo diga "SaaS (administracion)" y que leer_costeo lea el precio nuevo. Los eventos CONFIG_
se cuentan como NUEVOS (despues menos antes): asi la prueba sirve tambien sobre una copia que ya
traiga eventos CONFIG_ de otras pantallas (antes exigia 7 en todo el libro).
"""
import os
import shutil
from pathlib import Path

ORIGINAL = Path("datos_prueba/maestro_prueba.xlsx")
TEMPORAL = Path("datos_prueba/_prueba_costeo.xlsx")

# Antes de importar conexion_drive: trabajar SOLO sobre la copia temporal
os.environ["ORIGEN_EXCEL"] = "local"
os.environ["EXCEL_LOCAL_PATH"] = str(TEMPORAL)

from openpyxl import load_workbook                                   # noqa: E402

from microservicios.grupo4_costeo import costeo                      # noqa: E402
from microservicios.grupo4_costeo.costeo import CosteoInvalido       # noqa: E402

resultados = []


def caso(numero, descripcion, funcion, *args, debe_fallar=False, estado=None, **kwargs):
    try:
        r = funcion(*args, **kwargs)
    except CosteoInvalido as error:
        ok = debe_fallar
        detalle = f"rechazado: {str(error)[:90]}"
    else:
        ok = (not debe_fallar) and (estado is None or r["estado"] == estado)
        cambios = ", ".join(f"{c['celda']} {c['antes']!r}->{c['despues']!r}" for c in r.get("cambios", []))
        detalle = f"{r['estado']} {cambios}" + (" | AVISO" if r.get("avisos") else "")
    resultados.append(ok)
    print(f"{'OK   ' if ok else 'FALLO'} {numero:>2}. {descripcion}\n         {detalle}")


def formulas(ruta):
    wb = load_workbook(ruta)
    salida = {}
    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for c in fila:
                v = c.value
                if hasattr(v, "text"):
                    v = v.text
                if isinstance(v, str) and v.startswith("="):
                    salida[(ws.title, c.coordinate)] = v
    return salida


def main():
    if not ORIGINAL.exists():
        print(f"No existe {ORIGINAL}")
        return
    shutil.copyfile(ORIGINAL, TEMPORAL)
    print(f"Copia temporal: {TEMPORAL} (el original no se toca)\n")

    caso(1, "Precio de referencia Baguette = 850", costeo.actualizar_precio_referencia, "P01", 850, estado="OK")
    caso(2, "Precio Baguette = 400 (bajo el costo): BLOQUEAR", costeo.actualizar_precio_referencia,
         "Baguette", 400, debe_fallar=True)
    caso(3, "Precio Baguette = 850 otra vez: SIN_CAMBIO", costeo.actualizar_precio_referencia,
         "P01", 850, estado="SIN_CAMBIO")
    caso(4, "Sueldo Panadero = 360000", costeo.actualizar_plaza, "Panadero", sueldo=360000, estado="OK")
    caso(5, "Plaza 'Otros' ambigua: RECHAZAR", costeo.actualizar_plaza, "Otros", personas=1, debe_fallar=True)
    caso(6, "Plaza nueva Repartidor (1 persona, 300000)", costeo.agregar_plaza, "Repartidor", 1, 300000,
         estado="OK")
    caso(7, "Plaza Repartidor repetida: RECHAZAR", costeo.agregar_plaza, "Repartidor", 1, 300000,
         debe_fallar=True)
    caso(8, "Electricidad marzo = 210000", costeo.actualizar_gasto, "Electricidad", 3, 210000, estado="OK")
    caso(9, "Rubro 'Planilla' (es formula): RECHAZAR", costeo.actualizar_gasto, "Planilla", 1, 5,
         debe_fallar=True)
    caso(10, "Patente en junio (ya hay en enero): AVISO y se guarda", costeo.actualizar_gasto,
         "Patente", "junio", 50000, estado="OK")
    caso(11, "Mes 13: RECHAZAR", costeo.actualizar_gasto, "Internet", 13, 22000, debe_fallar=True)
    caso(12, "Activo nuevo Refrigerador (1 x 950000, 8 años)", costeo.agregar_activo,
         "Refrigerador", 1, 950000, 8, estado="OK")
    caso(13, "Dar de baja Espátula (cantidad 0)", costeo.actualizar_activo, "Espátula", cantidad=0,
         estado="OK")
    caso(14, "Activo inexistente: RECHAZAR", costeo.actualizar_activo, "Nave espacial", cantidad=1,
         debe_fallar=True)
    caso(15, "Precio de referencia CupCake = 345 con origen 'SaaS (administracion)' (#78)",
         costeo.actualizar_precio_referencia, "Pa01", 345, origen="SaaS (administracion)", estado="OK")

    print("\n=== Revision final del libro ===")
    antes, despues = formulas(ORIGINAL), formulas(TEMPORAL)
    distintas = [k for k in antes if despues.get(k) != antes[k]] + [k for k in despues if k not in antes]
    print(f"Formulas del original: {len(antes):,}   iguales despues de las pruebas: "
          f"{len(antes) - len([k for k in antes if despues.get(k) != antes[k]]):,}")
    resultados.append(not distintas)
    print(("OK    " if not distintas else "FALLO ") + f"Formulas cambiadas: {len(distintas)} {distintas[:5]}")

    v = load_workbook(TEMPORAL, data_only=True)
    va, fcp, pl = v["valor activos "], v["flujo caja proyecto"], v["planilla"]
    # Eventos CONFIG_ de costeo NUEVOS: los de la copia temporal que no estaban en el original.
    def eventos_config(libro):
        return [(fila[0].value, fila[3].value, fila[4].value) for fila in libro["RegistroEventos"].iter_rows(min_row=5)
                if str(fila[3].value or "").startswith("CONFIG_")]
    de_antes = {e[0] for e in eventos_config(load_workbook(ORIGINAL, data_only=True))}
    nuevos = [e for e in eventos_config(v) if e[0] not in de_antes]
    origenes = [e[2] for e in nuevos]
    cupcake = next(p for p in costeo.leer_costeo(v) if p["codigo"] == "Pa01")
    chequeos = [
        ("Total activos D36 = 7,817,500 (6,876,500 + 950,000 - 9,000)", va["D36"].value == 7817500),
        ("flujo caja proyecto D23 = D36", fcp["D23"].value == va["D36"].value),
        ("Depreciacion Refrigerador F14 = 118,750", va["F14"].value == 118750),
        ("Planilla: sueldo total Repartidor D13 = 300,000", pl["D13"].value == 300000),
        ("RegistroEventos: 8 eventos CONFIG_ nuevos (7 de antes + el del caso 15)", len(nuevos) == 8),
        ("Origen: los 7 primeros 'Robot Python' y el del caso 15 'SaaS (administracion)' (#78)",
         origenes[:7] == ["Robot Python"] * 7 and origenes[7:] == ["SaaS (administracion)"]),
        ("leer_costeo lee el precio de referencia nuevo de CupCake (345) y su utilidad",
         cupcake["precio_referencia"] == 345
         and round(cupcake["utilidad"], 2) == round(345 - cupcake["costo_unitario"], 2)),
    ]
    for texto, ok in chequeos:
        resultados.append(ok)
        print(("OK    " if ok else "FALLO ") + texto)

    print(f"\nRESULTADO: {sum(resultados)} de {len(resultados)} pruebas OK")
    print("TODO BIEN." if all(resultados) else "HAY PRUEBAS QUE FALLARON: revisar arriba.")


if __name__ == "__main__":
    main()
