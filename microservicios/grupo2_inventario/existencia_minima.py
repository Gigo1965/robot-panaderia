"""Microservicio: existencia minima por producto (Grupo 2, Decisiones Pendientes #63).

Mapa de celdas 'Mapa_Control Inventario': columna G "Stock (minimo): Umbral minimo de
existencia antes de reabastecer. Dato manual. Valor fijo definido por el usuario (hoy 10
para todos)". La columna H la calcula el Excel: IF(F<G, "REABASTECER", "OK").

Hallazgo presentado al usuario (4-oct-2026): el mapa describe G5:G11 (extracto de 7
productos) y el Excel tiene TablaControlInventario A4:J22 con 18 productos. Con el minimo
de 10 para todos (unidades, sacos y cajas por igual) casi toda la materia prima sale en
REABASTECER. Esta funcion permite que la administracion ponga el minimo de cada producto.

Decisiones del usuario (4-oct-2026):
  * La funcion cubre TODA la tabla (hoy G5:G22): ubica el producto por su nombre en la
    columna 'Producto Maestro' y la celda por el encabezado 'Stock' (no por la letra).
  * Se acepta un numero de 0 en adelante, con hasta 2 decimales (la materia prima se mide
    en fracciones de saco o de caja: 0.5 saco de harina es un minimo valido).
  * Lo cambia solo la administracion (el SaaS exige ese permiso).

Reglas del proyecto que se respetan:
  * Es CONFIGURACION (como descuento_general.py y costeo.py): el robot actualiza el valor
    (UPDATE) y deja un evento en RegistroEventos con el VALOR ANTERIOR -> VALOR NUEVO.
  * Solo se escribe la celda del minimo. Las existencias y el estado los calcula SIEMPRE
    el Excel con sus formulas (Regla #27); nunca se escribe sobre una formula.
  * Los nombres de producto son LLAVE (Regla #16): se comparan exactos (sin espacios a los
    lados). Un nombre que no esta en la tabla se rechaza y no se escribe nada.
  * Si el valor nuevo es igual al actual, no se sube nada (estado SIN_CAMBIO).

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo2_inventario.existencia_minima --local --ver
    py -m microservicios.grupo2_inventario.existencia_minima --local "HARINA FLORES 25 KG FUERTE" 2
"""
import math
import os
import sys

# --local: igual que descuento_general.py y las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import editar_excel, leer_excel                                   # noqa: E402
from conexion_drive.tablas import columnas_de_tabla, leer_tabla, obtener_tabla, registrar_evento   # noqa: E402

HOJA = "Control de Inventario"
TABLA = "TablaControlInventario"
COL_PRODUCTO = "Producto Maestro"
COL_MINIMO = "Stock"                       # columna G: existencia minima (dato manual segun el mapa)
COL_EXISTENCIA = "Existencia Actual"       # columna F (formula)
COL_ESTADO = "Estado"                      # columna H (formula)
MAXIMO = 99999                             # tope de seguridad contra un numero mal digitado
TIPO_EVENTO = "CONFIG_EXISTENCIA_MINIMA"


class MinimoInvalido(Exception):
    """Dato rechazado por una validacion: no se escribio nada."""


class ProductoNoEstaEnInventario(MinimoInvalido):
    """El nombre no esta en la columna 'Producto Maestro' de Control de Inventario."""


def normalizar_minimo(valor):
    """'2', '2,5', '0.5' o 2 -> numero de 0 en adelante con 2 decimales. Lo demas se rechaza."""
    if isinstance(valor, bool):
        raise MinimoInvalido(f"{valor!r} no es una cantidad.")
    texto = str(valor if valor is not None else "").strip().replace(",", ".")
    try:
        numero = float(texto)
    except ValueError:
        raise MinimoInvalido(f"{valor!r} no es una cantidad. Escriba un numero, por ejemplo 2 o 0.5.") from None
    if math.isnan(numero) or math.isinf(numero):
        raise MinimoInvalido(f"{valor!r} no es una cantidad. Escriba un numero, por ejemplo 2 o 0.5.")
    numero = round(numero, 2)
    if not 0 <= numero <= MAXIMO:
        raise MinimoInvalido(
            f"La existencia minima debe estar entre 0 y {MAXIMO} (llego {numero:g}). No se escribio nada.")
    return int(numero) if numero == int(numero) else numero


def _numero(valor):
    return valor if isinstance(valor, (int, float)) and not isinstance(valor, bool) else None


def leer_existencias(wb_valores=None):
    """Las filas de 'Control de Inventario' tal como las calcula el Excel (solo lectura).

    Devuelve una lista de diccionarios, en el orden de la hoja:
      {'fila', 'producto', 'entrada_compra', 'entrada_produccion', 'ventas_despacho',
       'ventas_digitales', 'salidas_combos', 'consumo_produccion', 'existencia', 'minimo', 'estado'}
    """
    wb_valores = wb_valores or leer_excel()
    ws, tabla = obtener_tabla(wb_valores, HOJA, TABLA)
    filas = []
    for fila in leer_tabla(ws, tabla):
        producto = str(fila.get(COL_PRODUCTO) or "").strip()
        if not producto:
            continue
        filas.append({
            "fila": fila["_fila"], "producto": producto,
            "entrada_compra": _numero(fila.get("Entrada Compra")),
            "entrada_produccion": _numero(fila.get("Entrada Producción")),
            "ventas_despacho": _numero(fila.get("Ventas Despacho")),
            "ventas_digitales": _numero(fila.get("Ventas Digitales")),
            "salidas_combos": _numero(fila.get("Salidas en Combos")),
            "consumo_produccion": _numero(fila.get("Consumo en Producción")),
            "existencia": _numero(fila.get(COL_EXISTENCIA)),
            "minimo": _numero(fila.get(COL_MINIMO)),
            "estado": str(fila.get(COL_ESTADO) or "").strip(),
        })
    return filas


def _buscar_fila(ws, tabla, producto):
    buscado = str(producto or "").strip()
    if not buscado:
        raise ProductoNoEstaEnInventario("Falta el producto. No se escribio nada.")
    for fila in leer_tabla(ws, tabla):
        if str(fila.get(COL_PRODUCTO) or "").strip() == buscado:
            return fila
    raise ProductoNoEstaEnInventario(
        f"'{buscado}' no esta en la columna '{COL_PRODUCTO}' de '{HOJA}'. El nombre debe ser exacto. "
        "No se escribio nada.")


def actualizar_existencia_minima(producto, minimo, origen="SaaS"):
    """Cambia la existencia minima (columna 'Stock') de UN producto de Control de Inventario.

    producto: nombre exacto de la columna 'Producto Maestro' (p. ej. 'HARINA FLORES 25 KG FUERTE').
    minimo:   numero de 0 en adelante, hasta 2 decimales ('2', '0,5', 10...).
    """
    minimo = normalizar_minimo(minimo)

    def _editar(wb, wb_valores):
        ws, tabla = obtener_tabla(wb, HOJA, TABLA)
        fila = _buscar_fila(ws, tabla, producto)
        nombre = str(fila[COL_PRODUCTO]).strip()
        celda = ws[f"{columnas_de_tabla(tabla)[COL_MINIMO]}{fila['_fila']}"]
        anterior = celda.value
        if isinstance(anterior, str) and anterior.startswith("="):
            raise MinimoInvalido(f"'{HOJA}'!{celda.coordinate} es una formula: el robot no la sobrescribe.")
        if _numero(anterior) is not None and round(anterior, 2) == minimo:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", "celda": celda.coordinate,
                    "producto": nombre, "minimo": minimo}

        # Solo para dejarlo anotado en el evento: lo que el Excel tenia calculado ANTES del cambio.
        ws_v, tabla_v = obtener_tabla(wb_valores, HOJA, TABLA)
        antes = next((f for f in leer_tabla(ws_v, tabla_v) if f["_fila"] == fila["_fila"]), {})
        existencia = _numero(antes.get(COL_EXISTENCIA))
        nota = f" | existencia {round(existencia, 4):g}, estado antes {antes.get(COL_ESTADO)}" if existencia is not None else ""

        celda.value = minimo
        evento = registrar_evento(
            wb, modulo="Control de Inventario", tipo_evento=TIPO_EVENTO, entidad=HOJA, id_entidad=celda.coordinate,
            accion="UPDATE", estado="OK", origen=origen, valor_referencia=minimo,
            detalle=f"'{HOJA}'!{celda.coordinate} existencia minima de {nombre}: {anterior!r} -> {minimo!r}{nota}",
        )
        return {"hubo_cambios": True, "estado": "OK", "celda": celda.coordinate, "evento": evento,
                "producto": nombre, "antes": anterior, "despues": minimo, "minimo": minimo}

    return editar_excel(_editar, descripcion=f"Existencia minima de {str(producto).strip()[:40]}: {minimo:g}")


if __name__ == "__main__":
    argumentos = [a for a in sys.argv[1:] if a != "--local"]
    if argumentos == ["--ver"]:
        for f in leer_existencias():
            existencia = "" if f["existencia"] is None else f"{round(f['existencia'], 2):g}"
            print(f"  fila {f['fila']:>2}  {f['producto'][:45]:<45}  existencia {existencia:>8}  "
                  f"minimo {f['minimo']!s:>6}  {f['estado']}")
        sys.exit(0)
    if len(argumentos) != 2:
        sys.exit('Uso: py -m microservicios.grupo2_inventario.existencia_minima [--local] --ver\n'
                 '     py -m microservicios.grupo2_inventario.existencia_minima [--local] "<producto>" <minimo>')
    try:
        print(actualizar_existencia_minima(argumentos[0], argumentos[1], origen="Terminal"))
    except MinimoInvalido as error:
        sys.exit(f"RECHAZADO: {error}")
