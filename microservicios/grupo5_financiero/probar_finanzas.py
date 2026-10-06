"""Prueba completa de finanzas.py sobre una COPIA TEMPORAL del Excel de prueba.

No toca datos_prueba/maestro_prueba.xlsx ni Google Drive: copia el Excel de prueba a
datos_prueba/_prueba_finanzas.xlsx (se reemplaza en cada corrida) y ahi ejecuta 26 casos:
9 cambios validos, 2 repetidos (SIN_CAMBIO) y 15 que DEBEN ser rechazados. Despues revisa lo que
calculo el Excel (egresos del dia, saldo de caja), los eventos, que la fila 26 de 'flujo caja anual'
siga vacia y que TODAS las formulas del libro quedaron identicas al original.

La prueba no da por hecho ningun valor del Excel: compara contra lo que lee al empezar.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo5_financiero.probar_finanzas
Tarda cerca de 2 minutos: cada cambio valido se recalcula con LibreOffice, igual que en produccion.
"""
import os
import shutil
from datetime import timedelta
from pathlib import Path

ORIGINAL = Path("datos_prueba/maestro_prueba.xlsx")
TEMPORAL = Path("datos_prueba/_prueba_finanzas.xlsx")

# Antes de importar conexion_drive: trabajar SOLO sobre la copia temporal
os.environ["ORIGEN_EXCEL"] = "local"
os.environ["EXCEL_LOCAL_PATH"] = str(TEMPORAL)

from openpyxl import load_workbook                                       # noqa: E402

from conexion_drive.recalculo import buscar_errores                     # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla              # noqa: E402
from conexion_drive.tiempo import ahora_local                            # noqa: E402
from microservicios.grupo5_financiero import finanzas                    # noqa: E402
from microservicios.grupo5_financiero.finanzas import FinanzasInvalido   # noqa: E402

resultados = []


def caso(numero, descripcion, funcion, *args, debe_fallar=False, estado=None, **kwargs):
    try:
        r = funcion(*args, **kwargs)
    except FinanzasInvalido as error:
        ok = debe_fallar
        detalle = f"rechazado: {str(error)[:100]}"
    else:
        ok = (not debe_fallar) and (estado is None or r["estado"] == estado)
        detalle = f"{r['estado']} {r['celda']} {r['antes']!r} -> {r['despues']!r}"
    resultados.append(ok)
    print(f"{'OK   ' if ok else 'FALLO'} {numero:>2}. {descripcion}\n         {detalle}")


def revisar(condicion, texto):
    resultados.append(bool(condicion))
    print(f"{'OK   ' if condicion else 'FALLO'}     {texto}")


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


def valores():
    return load_workbook(TEMPORAL, data_only=True)


def eventos(wb):
    ws, tabla = obtener_tabla(wb, "RegistroEventos", "tblRegistroEventos")
    return leer_tabla(ws, tabla)


def dia_de(leido, fecha):
    return next(d for d in leido["dias"] if d["fecha"] == fecha)


def main():
    if not ORIGINAL.exists():
        print(f"No existe {ORIGINAL}")
        return 1
    shutil.copyfile(ORIGINAL, TEMPORAL)
    print(f"Copia temporal: {TEMPORAL} (el original no se toca)\n")

    hoy = ahora_local().date()
    ayer = hoy - timedelta(days=1)
    wb0 = valores()
    inicio = finanzas.leer_flujo_diario(wb0, hoy=hoy, dias=400)
    primera = inicio["primera"]
    eventos_antes = len(eventos(wb0))
    caja_hoy, caja_ayer, caja_primera = (dia_de(inicio, f)["caja_chica"] or 0 for f in (hoy, ayer, primera))
    total_caja = next(r["total"] for r in inicio["resumen"] if "caja chica" in r["concepto"].casefold())
    anual0 = finanzas.leer_flujo_anual(wb0)
    marzo0 = anual0["salidas"]["meses"][2] or 0
    print(f"Hoy {hoy}; primer dia de la hoja {primera}; saldo de caja al empezar {inicio['saldo']:,.2f}\n")

    print("--- Caja chica del dia ('Flujo caja diario' columna G) ---")
    caso(1, "Caja chica de hoy = 3500 (vales)", finanzas.registrar_caja_chica, hoy, caja_hoy + 3500, "Vales de prueba",
         hoy=hoy, estado="OK")
    leido = finanzas.leer_flujo_diario(valores(), hoy=hoy)
    fila_hoy = dia_de(leido, hoy)
    revisar(fila_hoy["calculado"] and fila_hoy["caja_chica"] == caja_hoy + 3500
            and round(fila_hoy["egresos"] - (fila_hoy["materia_prima"] + fila_hoy["compras_externas"]), 2) == caja_hoy + 3500,
            f"el Excel sumo la caja chica a los egresos de hoy (fila {fila_hoy['fila']}: egresos {fila_hoy['egresos']:,.2f})")
    revisar(round(inicio["saldo"] - leido["saldo"], 2) == 3500,
            f"el saldo de caja ({leido['celda_saldo']}) bajo 3,500: {inicio['saldo']:,.2f} -> {leido['saldo']:,.2f}")
    saldo_ayer = dia_de(leido, ayer)["saldo"]          # ya calculado; la caja chica de hoy no lo mueve
    caso(2, "El mismo total otra vez: SIN_CAMBIO", finanzas.registrar_caja_chica, hoy.isoformat(), caja_hoy + 3500,
         "Vales de prueba", hoy=hoy, estado="SIN_CAMBIO")
    caso(3, "Total nuevo de hoy = 5000.50: REEMPLAZA al anterior", finanzas.registrar_caja_chica,
         hoy.strftime("%d/%m/%Y"), f"{caja_hoy + 5000.5:.2f}".replace(".", ","), "Vales y un flete", hoy=hoy, estado="OK")
    leido = finanzas.leer_flujo_diario(valores(), hoy=hoy)
    revisar(dia_de(leido, hoy)["caja_chica"] == caja_hoy + 5000.5 and round(inicio["saldo"] - leido["saldo"], 2) == 5000.5,
            "queda un solo total en el dia (5,000.50) y el saldo baja 5,000.50 respecto al inicio")
    caso(4, "Caja chica de ayer = 1200 (un dia que ya paso)", finanzas.registrar_caja_chica, ayer, caja_ayer + 1200,
         "Compra de bolsas", hoy=hoy, estado="OK")
    caso(5, "Caja chica del primer dia de la hoja = 800", finanzas.registrar_caja_chica, primera, caja_primera + 800,
         "Vale de arranque", hoy=hoy, estado="OK")
    leido = finanzas.leer_flujo_diario(valores(), hoy=hoy, dias=400)
    total_ahora = next(r["total"] for r in leido["resumen"] if "caja chica" in r["concepto"].casefold())
    revisar(round(total_ahora - total_caja, 2) == 7000.5 and round(inicio["saldo"] - leido["saldo"], 2) == 7000.5
            and round(saldo_ayer - dia_de(leido, ayer)["saldo"], 2) == 2000,
            "tres dias con caja chica: el total del periodo sube 7,000.50, el saldo de ayer baja 2,000 (800 + 1,200) "
            "y el de hoy 7,000.50")

    manana = hoy + timedelta(days=1)
    rechazos = (
        (6, "Fecha de mañana (futura)", (manana, 100, "x")),
        (7, "Fecha anterior al primer dia de la hoja", (primera - timedelta(days=1), 100, "x")),
        (8, "Fecha mal escrita", ("31-31-2026", 100, "x")),
        (9, "Monto negativo", (hoy, -1, "x")),
        (10, "Monto sobre el tope (500,001)", (hoy, 500001, "x")),
        (11, "Monto con letras", (hoy, "tres mil", "x")),
        (12, "Monto con separador de miles (1,500)", (hoy, "1,500", "x")),
        (13, "Motivo vacio", (hoy, 100, "   ")),
        (14, "Motivo que empieza con '='", (hoy, 100, "=SUM(A1)")),
        (15, "Motivo de mas de 120 letras", (hoy, 100, "x" * 121)),
    )
    antes = TEMPORAL.read_bytes()
    for numero, descripcion, argumentos in rechazos:
        caso(numero, f"{descripcion}: RECHAZAR", finanzas.registrar_caja_chica, *argumentos, hoy=hoy, debe_fallar=True)
    revisar(TEMPORAL.read_bytes() == antes, "los 10 rechazos no escribieron nada (el archivo quedo identico)")

    print("\n--- Salidas de efectivo del mes ('flujo caja anual' fila 19) ---")
    caso(16, "Salidas de efectivo de marzo = 25000", finanzas.registrar_salida_efectivo, "marzo", marzo0 + 25000,
         estado="OK")
    anual = finanzas.leer_flujo_anual(valores())
    otros_iguales = all(a == b for i, (a, b) in enumerate(zip(anual0["salidas"]["meses"], anual["salidas"]["meses"])) if i != 2)
    revisar(anual["salidas"]["meses"][2] == marzo0 + 25000 and otros_iguales,
            f"solo cambio la celda de marzo (fila {anual['salidas']['fila']}); los otros 11 meses quedaron igual")
    caso(17, "El mismo monto otra vez: SIN_CAMBIO", finanzas.registrar_salida_efectivo, 3, marzo0 + 25000,
         estado="SIN_CAMBIO")
    caso(18, "Salidas de efectivo de diciembre = 1000.75, con origen SaaS", finanzas.registrar_salida_efectivo, 12,
         "1000,75", origen="SaaS (administracion)", estado="OK")
    antes = TEMPORAL.read_bytes()
    for numero, descripcion, argumentos in ((19, "Mes 13", (13, 100)), (20, "Mes 'brumario'", ("brumario", 100)),
                                            (21, "Monto negativo", ("enero", -5)), (22, "Monto con letras", (1, "mucho")),
                                            (23, "Monto con separador de miles", (1, "25.000"))):
        caso(numero, f"{descripcion}: RECHAZAR", finanzas.registrar_salida_efectivo, *argumentos, debe_fallar=True)
    revisar(TEMPORAL.read_bytes() == antes, "los 5 rechazos no escribieron nada (el archivo quedo identico)")

    print("\n--- Eventos, formulas y regreso a los valores del inicio ---")
    wb = valores()
    nuevos = eventos(wb)[eventos_antes:]
    tipos = [e.get("TIPO_EVENTO") for e in nuevos]
    revisar(tipos == [finanzas.EVENTO_CAJA_CHICA] * 4 + [finanzas.EVENTO_SALIDAS] * 2,
            f"6 eventos nuevos, uno por cada cambio valido: {tipos}")
    revisar(all("->" in str(e.get("DETALLE")) for e in nuevos) and all("motivo:" in str(e.get("DETALLE")) for e in nuevos[:4]),
            "cada evento trae el valor anterior -> el nuevo; los de caja chica traen ademas el motivo")
    revisar([e.get("ORIGEN") for e in nuevos] == ["Robot Python"] * 5 + ["SaaS (administracion)"],
            "origen del evento: 'Robot Python' si nadie entrega otro; 'SaaS (administracion)' cuando lo entrega el SaaS")
    for evento in nuevos:
        print(f"         {evento.get('ID_EVENTO')} {evento.get('TIPO_EVENTO')}: {evento.get('DETALLE')}")

    caso(24, "Devolver la caja chica de hoy a su valor del inicio", finanzas.registrar_caja_chica, hoy, caja_hoy,
         "Regreso de la prueba", hoy=hoy, estado="OK")
    caso(25, "Devolver la de ayer", finanzas.registrar_caja_chica, ayer, caja_ayer, "Regreso de la prueba", hoy=hoy,
         estado="OK")
    caso(26, "Devolver la del primer dia", finanzas.registrar_caja_chica, primera, caja_primera, "Regreso de la prueba",
         hoy=hoy, estado="OK")
    leido = finanzas.leer_flujo_diario(valores(), hoy=hoy)
    revisar(round(leido["saldo"] - inicio["saldo"], 2) == 0,
            f"con la caja chica en sus valores del inicio, el saldo de caja vuelve a {inicio['saldo']:,.2f}")

    f_antes, f_despues = formulas(ORIGINAL), formulas(TEMPORAL)
    revisar(f_antes == f_despues, f"las {len(f_antes)} formulas del libro quedaron IDENTICAS al original")
    ws_anual = load_workbook(TEMPORAL)[finanzas.HOJA_ANUAL]
    fila_espacio = next(f["fila"] for f in finanzas.leer_flujo_anual(valores())["filas"]
                        if f["etiqueta"].casefold() == "gasto operacional") + 1
    revisar(all(ws_anual[f"{letra}{fila_espacio}"].value is None for letra in "A" + finanzas.COLUMNAS_MES),
            f"la fila {fila_espacio} de '{finanzas.HOJA_ANUAL}' (espacio de orden) sigue vacia: el robot no la toca")
    revisar(not buscar_errores(TEMPORAL.read_bytes()), "la copia temporal quedo con 0 errores de formula")

    print()
    fallos = resultados.count(False)
    if fallos:
        print(f"RESULTADO: {fallos} FALLO(S) de {len(resultados)} revisiones")
        return 1
    print(f"RESULTADO: TODO OK ({len(resultados)} de {len(resultados)} revisiones)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
