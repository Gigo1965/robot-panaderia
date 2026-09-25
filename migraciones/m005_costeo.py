"""Migracion m005 - Grupo 4 (Costeo): cadena de precios de proveedores + filas de reserva en activos.

Diseno aprobado por el usuario el 25-sep-2026.

Regla del usuario que la origina: cuando entra un producto nuevo (o cambia su precio) se revisa
que la secuencia de formulas internas del Excel tome el dato que llega de la API de catalogo
(TABLA CRUCE MAESTRA MP / TABLA CRUCE MAESTRA). La revision celda por celda encontro:

  A) Formulas de costo (hojas 'costo cupcakes' y 'costo pan baguette'):
     A1 cupcakes C13 margarina: tomaba el costo de 1 gramo (H17) -> porcion de 10 g (H18)
     A2 cupcakes C17 levadura:  tomaba el costo de 1 gramo (H15) -> porcion de 1.5 g (H16)
        (A1 y A2 reemplazan una decision anterior anotada como "diseno intencional" en
        DiccionarioDatos N180/N183; el usuario la reviso el 25-sep-2026: "hay que corregir
        para que tenga logica dentro de costos".)
     A3 cupcakes C31 electricidad: MONTH($C$6) leia la HORA (siempre enero) -> MONTH($C$7) (fecha)
     A4 cupcakes H21 leche: precio de VENTA (TCM G10) -> precio del proveedor (TCM D10).
        Resuelve Decisiones Pendientes #3.
     A5 cupcakes G5 peso de harina: 0.041 kg -> 0.025 kg (la receta usa 25 g; solo informativo)
     A6 baguette C16 empaque: 5 colones escritos a mano -> bolsa kraft real (H22, TCM MP E17).
        1 bolsa por baguette, igual que el Consumo de Materia Prima.
     A7 cupcakes G17 etiqueta 'Mantequilla 1gr' -> 'Margarina 1gr' (el insumo es margarina).
  B) valor activos: 20 filas de reserva dentro de Tabla12 (filas 14 a 33), con formulas que
     no dan error si la fila esta vacia. Los totales bajan de las filas 14-16 a las 34-36
     (mismas formulas, mismo formato). Tabla12 pasa de A2:H13 a A2:H33. Asi el robot agrega
     un activo escribiendo SOLO valores (Regla #7), sin insertar filas.
  C) flujo caja proyecto D23 (inversion en activos) sigue al total: D16 -> D36.
  D) DiccionarioDatos: notas '|| m005' (se AGREGA texto al final de la columna N, no se borra).

Seguridad (protocolo del proyecto, igual que la m004):
  * Antes de escribir revisa TODAS las celdas. Si alguna no tiene el valor "antes" esperado
    (ni el "despues"), se detiene SIN cambiar nada.
  * Si una celda ya tiene el valor "despues", la salta (se puede correr 2 veces).
  * Despues de escribir, vuelve a leer cada celda y confirma el valor.
  * Solo las celdas marcadas como FORMULA pueden empezar con "="; ningun TEXTO empieza con "=".
  * El evento MIGRACION guarda cada valor anterior -> valor nuevo (no se pierde informacion).

Uso (desde la carpeta ROBOT PANADERO):
    py -m migraciones.m005_costeo --local            (copia datos_prueba)
    py -m migraciones.m005_costeo                    (Excel real en Drive)
    py -m migraciones.m005_costeo --archivo "ruta.xlsx"   (un archivo suelto)
"""
import os
import sys
from copy import copy

# --local: igual que m001-m004. Debe ir ANTES de importar conexion_drive.config.
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

ID_MIGRACION = "m005_costeo"
FECHA = "25-sep-2026"

CUP = "costo cupcakes"
BAG = "costo pan baguette"
ACT = "valor activos "          # OJO: el nombre real de la hoja lleva un espacio al final
FCP = "flujo caja proyecto"
DIC = "DiccionarioDatos"

F, T = "FORMULA", "TEXTO"       # tipo de cada cambio

# ---------------------------------------------------------------- Bloque A (formulas de costo)
BLOQUE_A = [
    (CUP, "C13", "=H17", "=H18", F),
    (CUP, "C17", "=H15", "=H16", F),
    (CUP, "C31", "=INDEX('gasto operacional'!$B$8:$M$8, MONTH($C$6))",
                 "=INDEX('gasto operacional'!$B$8:$M$8, MONTH($C$7))", F),
    (CUP, "H21", "='TABLA CRUCE MAESTRA'!$G$10", "='TABLA CRUCE MAESTRA'!$D$10", F),
    (CUP, "G5", 0.041, 0.025, T),
    (BAG, "C16", 5, "=H22", F),
    (CUP, "G17", "Mantequilla 1gr", "Margarina 1gr", T),
]

# ---------------------------------------------------------------- Bloque B (valor activos)
PRIMERA_RESERVA, ULTIMA_RESERVA = 14, 33
TABLA_ANTES, TABLA_DESPUES = "A2:H13", "A2:H33"

# Totales: (columna, fila antes, fila despues, contenido). Se mueven 20 filas hacia abajo.
TOTALES = [
    ("A", 14, 34, "total depreciacion"),
    ("F", 14, 34, "=SUM(Tabla12[DEPRECIACION ANUAL])"),
    ("G", 14, 34, "=SUM(Tabla12[DEPRECIACION MENSUAL])"),
    ("A", 15, 35, "valor rescate a 12 meses"),
    ("H", 15, 35, "=SUM(Tabla12[Valor registrado a 12 meses])"),
    ("A", 16, 36, "Total activos"),
    ("D", 16, 36, "=SUM(Tabla12[COSTO])"),
]


def _formulas_reserva(fila):
    """Formulas de una fila de reserva: si la fila esta vacia muestran "" (no dan error)."""
    r = fila
    return {
        "D": f'=IF(A{r}="","",B{r}*C{r})',
        "F": f'=IF(OR(D{r}="",E{r}="",E{r}=0),"",D{r}/E{r})',
        "G": f'=IF(F{r}="","",ROUND(F{r}/12,0))',
        "H": f'=IF(OR(D{r}="",F{r}=""),"",D{r}-F{r})',
    }


def _estado_activos(momento):
    """Contenido esperado de A14:H36 'antes' o 'despues' de la migracion."""
    celdas = {f"{c}{r}": None for r in range(14, 37) for c in "ABCDEFGH"}
    if momento == "antes":
        for col, fila_antes, _, valor in TOTALES:
            celdas[f"{col}{fila_antes}"] = valor
    else:
        for fila in range(PRIMERA_RESERVA, ULTIMA_RESERVA + 1):
            for col, formula in _formulas_reserva(fila).items():
                celdas[f"{col}{fila}"] = formula
        for col, _, fila_despues, valor in TOTALES:
            celdas[f"{col}{fila_despues}"] = valor
    return celdas


def _bloque_b():
    antes, despues = _estado_activos("antes"), _estado_activos("despues")
    cambios = []
    for celda in antes:
        if antes[celda] != despues[celda]:
            tipo = F if isinstance(despues[celda], str) and despues[celda].startswith("=") else T
            cambios.append((ACT, celda, antes[celda], despues[celda], tipo))
    return cambios


# ---------------------------------------------------------------- Bloque C (flujo caja proyecto)
BLOQUE_C = [(FCP, "D23", "='valor activos '!D16", "='valor activos '!D36", F)]

# ---------------------------------------------------------------- Bloque D (notas agregadas)
def _n(texto):
    return f" || m005 ({FECHA}): {texto}"


NOTAS_D = [
    (DIC, "N156", _n("CORREGIDO: C16 = H22 (bolsa kraft real, TCM MP E17 = precio del proveedor / 100 "
                     "unidades). 1 bolsa por baguette, igual que Consumo de Materia Prima. Antes: 5 fijo.")),
    (DIC, "N174", _n("CORREGIDO: G5 = 0.025 kg (25 g por unidad, igual que la receta I14). Antes: 0.041. "
                     "Solo informativo (G6 Total harina kilos); no cambia el costo.")),
    (DIC, "N180", _n("CORREGIDO: C13 = H18 (costo de la porcion de 10 g de margarina). Antes: H17 (1 gramo). "
                     "El usuario revisa y REEMPLAZA la decision anterior ('diseno intencional'): hay que "
                     "corregir para que tenga logica dentro de costos (el inventario ya descuenta 10 g por "
                     "cupcake y el baguette usa la porcion). La etiqueta G17 pasa de 'Mantequilla 1gr' a "
                     "'Margarina 1gr'.")),
    (DIC, "N182", _n("CORREGIDO: H21 = 'TABLA CRUCE MAESTRA'!D10 (precio del proveedor por litro). Antes: G10 "
                     "(precio de venta = costo x 1.05). Resuelve Decisiones Pendientes #3.")),
    (DIC, "N183", _n("CORREGIDO: C17 = H16 (costo de la porcion de 1.5 g de levadura). Antes: H15 (1 gramo). "
                     "El usuario revisa y REEMPLAZA la decision anterior ('diseno intencional'): hay que "
                     "corregir para que tenga logica dentro de costos (el inventario ya descuenta 1.5 g por "
                     "cupcake y el baguette usa la porcion).")),
    (DIC, "N191", _n("C31 = INDEX('gasto operacional'!B8:M8, MONTH(C7)): toma la electricidad del MES DE LA "
                     "FECHA de la hornada. Antes MONTH(C6) leia la hora (siempre enero). La electricidad es un "
                     "costo promedio mensual hasta produccion full estimada (decision del usuario).")),
    (DIC, "N96", _n("Tabla12 ampliada a A2:H33: filas 14 a 33 = 20 filas de reserva para activos nuevos "
                    "(el robot escribe solo nombre, cantidad, precio y vida util; baja = cantidad 0).")),
    (DIC, "N104", _n("MOVIDO: el total esta ahora en F34 (antes F14); misma formula.")),
    (DIC, "N105", _n("MOVIDO: el total esta ahora en G34 (antes G14); misma formula.")),
    (DIC, "N106", _n("MOVIDO: el total esta ahora en H35 (antes H15); misma formula.")),
    (DIC, "N107", _n("MOVIDO: el total esta ahora en D36 (antes D16); misma formula. "
                     "flujo caja proyecto D23 apunta a D36.")),
    (DIC, "N280", _n("D23 = 'valor activos '!D36 (el total de activos se movio de D16 a D36 por las filas "
                     "de reserva de Tabla12). Mismo valor.")),
]


class MigracionDetenida(Exception):
    """Una celda no tenia el valor esperado: no se escribio nada."""


def _lista_cambios(wb):
    """Arma la lista final (hoja, celda, antes, despues, tipo, bloque)."""
    cambios = [(h, c, a, d, t, "A") for h, c, a, d, t in BLOQUE_A]
    cambios += [(h, c, a, d, t, "B") for h, c, a, d, t in _bloque_b()]
    cambios += [(h, c, a, d, t, "C") for h, c, a, d, t in BLOQUE_C]
    for h, c, nota in NOTAS_D:
        actual = wb[h][c].value
        if isinstance(actual, str) and actual.endswith(nota):
            cambios.append((h, c, actual[: -len(nota)], actual, T, "D"))
        else:
            cambios.append((h, c, actual, (actual or "") + nota, T, "D"))
    return cambios


def _copiar_estilo(ws, origen, destino):
    """Copia el formato completo de una celda a otra."""
    if origen.has_style:
        destino.font = copy(origen.font)
        destino.fill = copy(origen.fill)
        destino.border = copy(origen.border)
        destino.alignment = copy(origen.alignment)
        destino.number_format = origen.number_format
        destino.protection = copy(origen.protection)


def _formato_activos(ws):
    """Filas 34-36 reciben el formato de los totales (filas 14-16); filas 14-33 el de la fila 13.

    Se guarda primero el formato de 14-16 (antes de sobrescribirlo) para no perderlo.
    """
    guardado = {}
    for fila in (14, 15, 16):
        guardado[fila] = [(copy(c.font), copy(c.fill), copy(c.border), copy(c.alignment),
                           c.number_format, copy(c.protection), c.has_style) for c in
                          (ws.cell(fila, col) for col in range(1, 9))]
    alturas = {f: ws.row_dimensions[f].height for f in (13, 14, 15, 16)}

    for fila_antes, fila_despues in ((14, 34), (15, 35), (16, 36)):
        for col in range(1, 9):
            font, fill, border, alignment, numfmt, protection, tiene = guardado[fila_antes][col - 1]
            if tiene:
                d = ws.cell(fila_despues, col)
                d.font, d.fill, d.border, d.alignment = font, fill, border, alignment
                d.number_format, d.protection = numfmt, protection
        if alturas[fila_antes]:
            ws.row_dimensions[fila_despues].height = alturas[fila_antes]

    for fila in range(PRIMERA_RESERVA, ULTIMA_RESERVA + 1):
        for col in range(1, 9):
            _copiar_estilo(ws, ws.cell(13, col), ws.cell(fila, col))
        if alturas[13]:
            ws.row_dimensions[fila].height = alturas[13]


def migrar(wb, wb_valores=None):
    cambios = _lista_cambios(wb)
    tabla = wb[ACT].tables["Tabla12"]

    # 1) Revisar TODO antes de escribir (incluido el rango de Tabla12)
    problemas, pendientes = [], []
    for hoja, celda, antes, despues, tipo, bloque in cambios:
        valor = wb[hoja][celda].value
        if valor == despues:
            continue
        if valor != antes:
            problemas.append(f"'{hoja}'!{celda}: se esperaba {antes!r} y tiene {valor!r}")
        else:
            pendientes.append((hoja, celda, antes, despues, tipo, bloque))
    if tabla.ref not in (TABLA_ANTES, TABLA_DESPUES):
        problemas.append(f"Tabla12: se esperaba {TABLA_ANTES} y tiene {tabla.ref}")
    if problemas:
        raise MigracionDetenida("NO SE CAMBIO NADA. Celdas distintas a lo esperado:\n  "
                                + "\n  ".join(problemas))

    # 2) Regla del proyecto: ningun TEXTO nuevo empieza con "=" (solo las FORMULAS)
    for hoja, celda, antes, despues, tipo, bloque in pendientes:
        if tipo == T and isinstance(despues, str) and despues.startswith("="):
            raise MigracionDetenida(f"Texto que empieza con '=' en '{hoja}'!{celda}. No se escribio nada.")

    tabla_pendiente = tabla.ref != TABLA_DESPUES
    if not pendientes and not tabla_pendiente:
        return {"hubo_cambios": False, "cambios": ["Ya estaba aplicada: 0 celdas pendientes."],
                "detalle": []}

    # 3) Formato de valor activos (ANTES de escribir, para guardar el de los totales)
    if any(b == "B" for *_, b in pendientes):
        _formato_activos(wb[ACT])

    # 4) Escribir valores y ampliar Tabla12
    for hoja, celda, antes, despues, tipo, bloque in pendientes:
        wb[hoja][celda] = despues
    if tabla_pendiente:
        tabla.ref = TABLA_DESPUES
        if tabla.autoFilter is not None:
            tabla.autoFilter.ref = TABLA_DESPUES

    # 5) Volver a leer y confirmar
    for hoja, celda, antes, despues, tipo, bloque in pendientes:
        if wb[hoja][celda].value != despues:
            raise MigracionDetenida(f"Verificacion fallo en '{hoja}'!{celda}")
    if tabla.ref != TABLA_DESPUES:
        raise MigracionDetenida("Verificacion fallo en el rango de Tabla12")

    por_bloque = {}
    for p in pendientes:
        por_bloque[p[5]] = por_bloque.get(p[5], 0) + 1
    resumen = [f"Bloque {b}: {n} celdas" for b, n in sorted(por_bloque.items())]
    if tabla_pendiente:
        resumen.append(f"Tabla12: {TABLA_ANTES} -> {TABLA_DESPUES}")

    # Valor anterior -> valor nuevo de las celdas de formula de costo (no se pierde informacion)
    antes_despues = " ; ".join(f"'{h}'!{c}: {a} -> {d}" for h, c, a, d, t, b in pendientes if b in "AC")

    from conexion_drive.tablas import registrar_evento
    registrar_evento(
        wb, modulo="Sistema", tipo_evento="MIGRACION", entidad="Libro Maestro",
        id_entidad=ID_MIGRACION, accion="Costeo G4 (cadena de precios + activos)", estado="OK",
        detalle=("A) formulas de costo conectadas a la API de catalogo | B) valor activos: 20 filas "
                 "de reserva (14-33), totales a 34-36, Tabla12 A2:H33 | C) flujo caja proyecto D23 -> D36 | "
                 f"D) notas '|| m005' en DiccionarioDatos | Total: {len(pendientes)} celdas | "
                 f"Antes -> despues: {antes_despues}"),
    )
    return {"hubo_cambios": True, "cambios": resumen, "detalle": pendientes}


def _aplicar_a_archivo(ruta):
    """Aplica la migracion a un .xlsx suelto y lo recalcula con LibreOffice."""
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
            resultado = editar_excel(migrar, descripcion="Migracion m005_costeo")
    except MigracionDetenida as error:
        print(str(error))
        sys.exit(1)
    for linea in resultado["cambios"]:
        print("  " + linea)
    print("LISTO." if resultado["hubo_cambios"] else "Ya estaba aplicada: no se cambio nada.")
