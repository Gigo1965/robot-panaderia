"""Migracion m007 - Grupo 5 (Financiero): planilla administrativa, flujo caja proyecto (F1, F2) y F11.

Diseno aprobado por el usuario el 25-sep-2026. Cambia formulas de 'gasto operacional' (fila 9),
'flujo caja proyecto' (años 2026-2030) y 'Punto de Equilibrio' (B23):

  A) Concepto del usuario (contabilidad de costos): costo de produccion = materia prima + mano de
     obra directa (panadero y ayudante) + costos indirectos; gasto administrativo = administrador,
     contador, miscelaneos. 'gasto operacional' fila 9 ("Planilla", B9:M9) pasa a sumar SOLO los
     salarios administrativos: planilla M14 (o M15 desde julio) MENOS la parte del sueldo de
     panadero y ayudante que ya esta en el costo unitario (80% baguette + 15% cupcake, celdas
     B23/B24 de cada hoja de costo). La hoja 'planilla' NO se toca: sigue siendo el total del staff.
     Cargas sociales, poliza y aguinaldo siguen completos en gasto operacional (decision del
     usuario). Resuelve Decisiones Pendientes #6. Se refleja solo en flujo caja proyecto B17,
     Punto de Equilibrio B15 y flujo caja anual fila 25 (todos leen 'gasto operacional').
  B) F2 - Cada año (2026-2030, columnas E a I) se calcula desde sus PROPIAS filas:
     venta (10) - costo (11) = utilidad bruta (13); - gasto (17) = antes de impuestos (18);
     - impuesto 7% (19) = utilidad neta (20). Se quita el x12 de venta y costo (las cantidades
     del Punto de Equilibrio ya son anuales). 2026 = año base; desde 2027 venta, costo y gasto
     crecen 3% (sin cambio). F22/G22 restan su propia cuota (F21/G21; mismo valor que E21).
  C) F1 - Utilidad neta (fila 20) = utilidad antes de impuestos - impuesto (antes no lo restaba).
  D) DiccionarioDatos: notas '|| m007' (se AGREGA texto, no se borra).
  E) F11 - Punto de Equilibrio B23: la utilidad esperada (20%) se calcula sobre TODAS las
     obligaciones del punto de equilibrio (gasto operacional + cuota del prestamo), no solo sobre
     el gasto operacional. Concepto del usuario: ganancia esperada = 20% sobre el punto de equilibrio.

NO cambia (decisiones del usuario): financiamiento (F7) y tasa de impuesto (F6), pendientes;
sin depreciacion (F5) ni valor de rescate (F8); costo = hornada estimada (F3, Pendiente #5).

Seguridad: igual que m005 (revisa TODO antes de escribir, se detiene sin cambiar nada si una
celda no tiene el valor esperado, se puede correr 2 veces, relee cada celda, evento MIGRACION
con cada formula anterior -> nueva).

Uso (desde la carpeta ROBOT PANADERO):
    py -m migraciones.m007_flujo_proyecto --local
    py -m migraciones.m007_flujo_proyecto
    py -m migraciones.m007_flujo_proyecto --archivo "ruta.xlsx"
"""
import os
import sys

if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

ID_MIGRACION = "m007_flujo_proyecto"
FECHA = "25-sep-2026"
FCP = "flujo caja proyecto"
DIC = "DiccionarioDatos"
AÑOS = "EFGHI"          # 2026 ... 2030

# ---------------------------------------------------------------- Bloque A (F4)
GOP = "gasto operacional"
PE = "Punto de Equilibrio"
_MOD = ("(('costo pan baguette'!$B$24+'costo cupcakes'!$B$24)*planilla!$D$8"
        "+('costo pan baguette'!$B$23+'costo cupcakes'!$B$23)*planilla!$D$9)")
BLOQUE_A = ([(GOP, f"{c}9", "=planilla!$M$14", f"=planilla!$M$14-{_MOD}") for c in "BCDEFG"]
            + [(GOP, f"{c}9", "=planilla!$M$15", f"=planilla!$M$15-{_MOD}*1.03") for c in "HIJKLM"])
BLOQUE_E = [(PE, "B23", "=B15*B22", "=(B15+B21)*B22")]

# ---------------------------------------------------------------- Bloques B y C (F2 y F1)
ANTES = {
    "E10": "=ROUND((B10+C10)*12,2)",
    "E11": "=ROUND(((B11*B8)+(C11*C8))*12,2)",
    "E13": "=ROUND(B13*1.03,2)", "F13": "=ROUND(E13*1.03,2)", "G13": "=ROUND(F13*1.03,2)",
    "H13": "=ROUND(G13*1.03,2)", "I13": "=ROUND(H13*1.03,2)",
    "E18": "=ROUND(B18,2)", "F18": "=ROUND(E18*1.03,2)", "G18": "=ROUND(F18*1.03,2)",
    "H18": "=ROUND(G18*1.03,2)", "I18": "=ROUND(H18*1.03,2)",
    "E19": "=ROUND(Tabla13[[#This Row],[Clase]],2)",
    "F19": "=ROUND(Tabla13[[#This Row],[2026]]*1.03,2)",
    "G19": "=ROUND(Tabla13[[#This Row],[2027]]*1.03,2)",
    "H19": "=ROUND(Tabla13[[#This Row],[2028]]*1.03,2)",
    "I19": "=ROUND(Tabla13[[#This Row],[2029]]*1.03,2)",
    "E20": "=ROUND(E18,2)", "F20": "=ROUND(F18,2)", "G20": "=ROUND(F20*1.03,2)",
    "H20": "=ROUND(H18,2)", "I20": "=ROUND(I18,2)",
    "F22": "=ROUND(F20-E21,2)", "G22": "=ROUND(G20-E21,2)",
}
DESPUES = {"E10": "=ROUND(B10+C10,2)", "E11": "=ROUND((B11*B8)+(C11*C8),2)",
           "F22": "=ROUND(F20-F21,2)", "G22": "=ROUND(G20-G21,2)"}
for c in AÑOS:
    DESPUES[f"{c}13"] = f"=ROUND({c}10-{c}11,2)"
    DESPUES[f"{c}18"] = f"=ROUND({c}13-{c}17,2)"
    DESPUES[f"{c}19"] = f"=ROUND(({c}18*7)/100,2)"
    DESPUES[f"{c}20"] = f"=ROUND({c}18-{c}19,2)"
BLOQUE_BC = [(FCP, celda, ANTES[celda], DESPUES[celda]) for celda in ANTES]


# ---------------------------------------------------------------- Bloque D (notas)
def _n(texto):
    return f" || m007 ({FECHA}): {texto}"


NOTAS_D = [
    (DIC, "N270", _n("E10 = B10+C10 (2026 = año base, SIN x12: las cantidades del Punto de Equilibrio ya son "
                     "anuales). Desde 2027 crece 3%.")),
    (DIC, "N271", _n("E11 = (B11xB8)+(C11xC8) (2026 = año base, SIN x12). Desde 2027 crece 3%.")),
    (DIC, "N272", _n("Años 2026-2030 (E13:I13) = venta - costo de cada año (antes: B13 x 1.03 encadenado).")),
    (DIC, "N215", _n("Fila 9 'Planilla' de gasto operacional = SOLO salarios administrativos: planilla M14 (M15 "
                     "desde julio) menos la parte del sueldo de panadero y ayudante que ya esta en el costo de "
                     "produccion (80% baguette + 15% cupcake). Concepto del usuario: mano de obra directa = costo "
                     "de produccion; administrador, contador, miscelaneos = gasto administrativo. La hoja "
                     "'planilla' sigue siendo el total del staff. Cargas, poliza y aguinaldo siguen completos "
                     "aqui. Resuelve Decisiones Pendientes #6.")),
    (DIC, "N274", _n("Años 2026-2030 (E18:I18) = utilidad bruta - gasto operacional de cada año.")),
    (DIC, "N275", _n("Años 2026-2030 (E19:I19) = 7% de la utilidad antes de impuestos de cada año (tasa F6 "
                     "pendiente de confirmar por el usuario).")),
    (DIC, "N276", _n("CORREGIDO años 2026-2030 (E20:I20) = utilidad antes de impuestos - impuesto. Antes la "
                     "utilidad neta era igual a la utilidad antes de impuestos (no restaba el impuesto).")),
    (DIC, "N279", _n("F22 y G22 restan su propia cuota (F21, G21); antes restaban E21 (mismo valor).")),
]


class MigracionDetenida(Exception):
    """Una celda no tenia el valor esperado: no se escribio nada."""


def _lista_cambios(wb):
    cambios = [(h, c, a, d, "A") for h, c, a, d in BLOQUE_A]
    cambios += [(h, c, a, d, "E") for h, c, a, d in BLOQUE_E]
    for h, c, a, d in BLOQUE_BC:
        bloque = "C" if c[1:] == "20" else "B"
        cambios.append((h, c, a, d, bloque))
    for h, c, nota in NOTAS_D:
        actual = wb[h][c].value
        if isinstance(actual, str) and actual.endswith(nota):
            cambios.append((h, c, actual[: -len(nota)], actual, "D"))
        else:
            cambios.append((h, c, actual, (actual or "") + nota, "D"))
    return cambios


def migrar(wb, wb_valores=None):
    cambios = _lista_cambios(wb)

    # 1) Revisar TODO antes de escribir
    problemas, pendientes = [], []
    for hoja, celda, antes, despues, bloque in cambios:
        valor = wb[hoja][celda].value
        if valor == despues:
            continue
        if valor != antes:
            problemas.append(f"'{hoja}'!{celda}: se esperaba {antes!r} y tiene {valor!r}")
        else:
            pendientes.append((hoja, celda, antes, despues, bloque))
    if problemas:
        raise MigracionDetenida("NO SE CAMBIO NADA. Celdas distintas a lo esperado:\n  "
                                + "\n  ".join(problemas))

    # 2) Ningun TEXTO (notas) empieza con "="; las formulas (bloques A, B, C) si
    for hoja, celda, antes, despues, bloque in pendientes:
        if bloque == "D" and isinstance(despues, str) and despues.startswith("="):
            raise MigracionDetenida(f"Texto que empieza con '=' en '{hoja}'!{celda}. No se escribio nada.")

    if not pendientes:
        return {"hubo_cambios": False, "cambios": ["Ya estaba aplicada: 0 celdas pendientes."],
                "detalle": []}

    # 3) Escribir
    for hoja, celda, antes, despues, bloque in pendientes:
        wb[hoja][celda] = despues

    # 4) Volver a leer y confirmar
    for hoja, celda, antes, despues, bloque in pendientes:
        if wb[hoja][celda].value != despues:
            raise MigracionDetenida(f"Verificacion fallo en '{hoja}'!{celda}")

    por_bloque = {}
    for p in pendientes:
        por_bloque[p[4]] = por_bloque.get(p[4], 0) + 1
    resumen = [f"Bloque {b}: {n} celdas" for b, n in sorted(por_bloque.items())]
    antes_despues = " ; ".join(f"'{h}'!{c}: {a} -> {d}" for h, c, a, d, b in pendientes if b != "D")

    from conexion_drive.tablas import registrar_evento
    registrar_evento(
        wb, modulo="Sistema", tipo_evento="MIGRACION", entidad="Libro Maestro",
        id_entidad=ID_MIGRACION, accion="Grupo 5: planilla administrativa, F1, F2, F11", estado="OK",
        detalle=("A) gasto operacional fila 9 = solo salarios administrativos | B) F2 años 2026-2030 desde sus "
                 "propias filas, sin x12 | C) F1 utilidad neta resta el impuesto | D) notas '|| m007' | E) F11 PE B23 = 20% sobre gasto + cuota | "
                 f"Total: {len(pendientes)} celdas | antes -> despues: {antes_despues}"),
    )
    return {"hubo_cambios": True, "cambios": resumen, "detalle": pendientes}


def _aplicar_a_archivo(ruta):
    import io
    from pathlib import Path
    import openpyxl
    from conexion_drive.recalculo import reparar_prefijos_xlfn, recalcular_bytes

    ruta = Path(ruta)
    contenido, _ = reparar_prefijos_xlfn(ruta.read_bytes())
    wb = openpyxl.load_workbook(io.BytesIO(contenido))
    resultado = migrar(wb, None)
    if resultado["hubo_cambios"]:
        salida = io.BytesIO()
        wb.save(salida)
        ruta.write_bytes(recalcular_bytes(salida.getvalue()))
    return resultado


if __name__ == "__main__":
    try:
        if "--archivo" in sys.argv:
            ruta = sys.argv[sys.argv.index("--archivo") + 1]
            print(f"Excel de destino: ARCHIVO {ruta}")
            resultado = _aplicar_a_archivo(ruta)
        else:
            from conexion_drive import config
            from conexion_drive.drive_excel import editar_excel
            print(f"Excel de destino: {'COPIA LOCAL ' + str(config.EXCEL_LOCAL_PATH) if config.ORIGEN_EXCEL == 'local' else 'GOOGLE DRIVE (real)'}")
            resultado = editar_excel(migrar, descripcion="Migracion m007_flujo_proyecto")
    except MigracionDetenida as error:
        print(str(error))
        sys.exit(1)
    for linea in resultado["cambios"]:
        print("  " + linea)
    print("LISTO." if resultado["hubo_cambios"] else "Ya estaba aplicada: no se cambio nada.")
