"""Ayudas genericas para trabajar con las Tablas de Excel (ListObjects) del maestro.

Reglas del proyecto que se respetan aqui:
- Las filas nuevas se AGREGAN al final de la tabla (INSERT). Nunca se usa
  insert_rows() (no ajusta el texto de las formulas existentes).
- El rango de la tabla se agranda para incluir la fila nueva, asi las
  referencias estructuradas (Tabla6[Total], etc.) la incluyen.
- La fila nueva copia el formato de la fila anterior (fecha, hora, bordes...).
- Las columnas se ubican por el NOMBRE del encabezado, no por la letra, para
  que el codigo no se rompa si algun dia se agrega una columna.
"""
from copy import copy

from openpyxl.utils import get_column_letter, range_boundaries

from .tiempo import ahora_local


def obtener_tabla(wb, nombre_hoja, nombre_tabla):
    ws = wb[nombre_hoja]
    if nombre_tabla not in ws.tables:
        raise KeyError(f"No existe la tabla '{nombre_tabla}' en la hoja '{nombre_hoja}'")
    return ws, ws.tables[nombre_tabla]


def columnas_de_tabla(tabla):
    """{'Fecha': 'C', 'Hora': 'D', ...} segun los encabezados de la tabla."""
    min_col, _, _, _ = range_boundaries(tabla.ref)
    return {col.name.strip(): get_column_letter(min_col + i) for i, col in enumerate(tabla.tableColumns)}


def filas_de_datos(ws, tabla):
    """Numeros de fila de datos de la tabla (sin el encabezado)."""
    _, min_row, _, max_row = range_boundaries(tabla.ref)
    return range(min_row + 1, max_row + 1)


def leer_tabla(ws, tabla):
    """Lista de diccionarios, una por fila con datos: [{'Fecha':..., 'Producto':...}, ...]"""
    columnas = columnas_de_tabla(tabla)
    registros = []
    for fila in filas_de_datos(ws, tabla):
        registro = {nombre: ws[f"{letra}{fila}"].value for nombre, letra in columnas.items()}
        if any(valor not in (None, "") for valor in registro.values()):
            registro["_fila"] = fila
            registros.append(registro)
    return registros


def _fila_vacia(ws, fila, letras):
    return all(ws[f"{letra}{fila}"].value in (None, "") for letra in letras)


def agregar_fila(ws, tabla, valores: dict, columnas_clave=None):
    """Agrega una fila al final de la tabla y devuelve su numero de fila.

    valores:        {'Fecha': ..., 'Producto': ...}  (nombres de encabezado)
    columnas_clave: encabezados que se revisan para saber si una fila esta vacia
                    (por defecto, las columnas que se van a escribir)
    """
    if tabla.totalsRowCount:
        raise NotImplementedError("Tablas con fila de totales no estan soportadas")

    columnas = columnas_de_tabla(tabla)
    desconocidas = set(valores) - set(columnas)
    if desconocidas:
        raise KeyError(f"Columnas que no existen en {tabla.name}: {sorted(desconocidas)}")

    min_col, min_row, max_col, max_row = range_boundaries(tabla.ref)
    letras_clave = [columnas[c] for c in (columnas_clave or valores.keys())]

    # Primera fila vacia dentro de la tabla, o la siguiente despues del final
    ultima_con_datos = min_row
    for fila in range(min_row + 1, max_row + 1):
        if not _fila_vacia(ws, fila, letras_clave):
            ultima_con_datos = fila
    fila_nueva = ultima_con_datos + 1

    # Copiar formato de la fila anterior (si es de datos)
    if fila_nueva - 1 > min_row:
        for col in range(min_col, max_col + 1):
            origen = ws.cell(fila_nueva - 1, col)
            destino = ws.cell(fila_nueva, col)
            if origen.has_style:
                destino.font = copy(origen.font)
                destino.border = copy(origen.border)
                destino.fill = copy(origen.fill)
                destino.number_format = origen.number_format
                destino.alignment = copy(origen.alignment)
                destino.protection = copy(origen.protection)

    for nombre, valor in valores.items():
        ws[f"{columnas[nombre]}{fila_nueva}"] = valor

    # Agrandar el rango de la tabla si la fila quedo afuera
    if fila_nueva > max_row:
        nuevo_rango = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{fila_nueva}"
        tabla.ref = nuevo_rango
        if tabla.autoFilter is not None:
            tabla.autoFilter.ref = nuevo_rango
    return fila_nueva


# ----------------------------------------------------------------------------
# RegistroEventos (historial append-only de todo lo que hace el robot)
# ----------------------------------------------------------------------------
def registrar_evento(wb, modulo, tipo_evento, entidad, id_entidad, accion, estado,
                     detalle="", valor_referencia="", origen="Robot Python",
                     usuario_proceso="robot", momento=None):
    ws, tabla = obtener_tabla(wb, "RegistroEventos", "tblRegistroEventos")
    numeros = []
    for registro in leer_tabla(ws, tabla):
        id_evento = str(registro.get("ID_EVENTO") or "")
        if id_evento.startswith("EVT-") and id_evento[4:].isdigit():
            numeros.append(int(id_evento[4:]))
    siguiente = max(numeros, default=0) + 1
    momento = momento or ahora_local()

    agregar_fila(ws, tabla, {
        "ID_EVENTO": f"EVT-{siguiente:06d}",
        "FECHA_HORA": momento.strftime("%Y-%m-%d %H:%M:%S"),
        "MÓDULO": modulo,
        "TIPO_EVENTO": tipo_evento,
        "ORIGEN": origen,
        "ENTIDAD": entidad,
        "ID_ENTIDAD": str(id_entidad),
        "ACCIÓN": accion,
        "ESTADO": estado,
        "USUARIO_PROCESO": usuario_proceso,
        "VALOR_REFERENCIA": valor_referencia,
        "DETALLE": detalle,
    }, columnas_clave=["ID_EVENTO"])
    return f"EVT-{siguiente:06d}"
