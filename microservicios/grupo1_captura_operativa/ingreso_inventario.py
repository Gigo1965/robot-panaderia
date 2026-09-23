"""Microservicio: Ingreso de Inventario (Grupo 1 - Captura operativa).

Sigue el 'Diseño del robot (4)' del libro de mapas V5:
  1. Quien recibe la compra (o el panadero, si es produccion propia) la
     registra en el SaaS: Fecha, Factura/Lote, Producto y Cantidad.
  2. El robot compara el producto contra la lista EXACTA de:
       - TABLA CRUCE MAESTRA    (Tipo 'Propio' = produccion, 'Tercero' = compra)
       - TABLA CRUCE MAESTRA MP (materia prima = compra)
  3. Si coincide: agrega la fila a TablaIngresos con el nombre EXACTO.
  (Antes de eso, mientras la persona escribe en el SaaS, buscar_productos_inventario()
   le muestra opciones con autocompletado: escribe "har" y elige "Harina".
   Asi nunca tiene que digitar el nombre exacto, aunque haya fila de clientes.)
  4. Si NO coincide: NO escribe nada y sugiere los nombres mas parecidos,
     para que la persona confirme en el SaaS. (Caso real: 'MEJORANTE STAR
     PAN PLUS 5 KILOS' debia ser '...3 KILOS'; si se deja pasar, el precio
     queda en 0 sin que nadie lo note.)

El robot NUNCA escribe precios: 'Precio unitario', 'Precio Unitario2', los
Totales y la Categoria son formulas que se calculan solas. Como las filas
nuevas no heredan formulas, el robot escribe el TEXTO de cada formula,
igual al de las filas existentes.

El producto decide la columna:
  - Tipo 'Propio' (Baguette, CupCake) -> 'Produccion propia' + 'Cantidad Produccion'
  - Todo lo demas                     -> 'Compra externa' + 'Cantidad Compras'

Duplicados (supuesto razonable, a confirmar): si ya existe una fila con la
misma Fecha + Factura + Producto + Cantidad, no se vuelve a escribir.
"""
from datetime import date, datetime
from difflib import SequenceMatcher

from conexion_drive import editar_excel
from conexion_drive.tablas import agregar_fila, leer_tabla, obtener_tabla, registrar_evento
from conexion_drive.tiempo import ahora_local

HOJA, TABLA = "Ingreso de Inventario", "TablaIngresos"

FORMULA_PRECIO_COMPRA = (
    '=IF(C{r}="","",IFERROR(INDEX(TablaProductos[Costo Unitario (Robot)], MATCH(C{r}, '
    'TablaProductos[Nombre tecnico completo], 0)), IFERROR(INDEX(TablaCruceMaestraMP[Costo Unitario '
    '(Scraping)], MATCH(C{r}, TablaCruceMaestraMP[Materia Prima], 0)), 0)))'
)
FORMULA_TOTAL_COMPRA = "=$D{r}*$E{r}"
FORMULA_PRECIO_PRODUCCION = (
    '=IF(G{r}="","",IFERROR(INDEX(TablaProductos[Costo Unitario (Robot)], MATCH(G{r}, '
    'TablaProductos[Nombre tecnico completo], 0)), 0))'
)
FORMULA_TOTAL_PRODUCCION = (
    "=TablaIngresos[[#This Row],[Cantidad Producción]]*TablaIngresos[[#This Row],[Precio Unitario2]]"
)
FORMULA_CATEGORIA = (
    '=IF(C{r}="","",IF(ISNUMBER(MATCH(C{r},TablaCruceMaestraMP[Materia Prima],0)),'
    '"Materia Prima","Compras Externas Netas"))'
)


class ProductoNoReconocido(Exception):
    """El nombre no coincide exacto. Trae 'sugerencias' para mostrar en el SaaS."""

    def __init__(self, escrito, sugerencias):
        self.escrito = escrito
        self.sugerencias = sugerencias
        texto = f"'{escrito}' no existe en TABLA CRUCE MAESTRA ni en TABLA CRUCE MAESTRA MP. No se escribio nada."
        if sugerencias:
            texto += " ¿Quiso decir: " + " | ".join(sugerencias) + "?"
        super().__init__(texto)


# ----------------------------------------------------------------------------
# Lista valida de productos para inventario
# ----------------------------------------------------------------------------
def _texto(valor):
    return "" if valor is None else str(valor).strip()


def leer_productos_inventario(wb_valores):
    """Todos los nombres validos para Ingreso de Inventario, con su tipo de movimiento."""
    productos = []
    ws, tabla = obtener_tabla(wb_valores, "TABLA CRUCE MAESTRA", "TablaProductos")
    for fila in leer_tabla(ws, tabla):
        if _texto(fila.get("Nombre tecnico completo")):
            productos.append({
                "nombre": fila["Nombre tecnico completo"],
                "codigo": _texto(fila.get("Codigo")),
                "visible": _texto(fila.get("Nombre visible")),
                "movimiento": "produccion" if _texto(fila.get("Tipo")).lower() == "propio" else "compra",
                "origen": "TABLA CRUCE MAESTRA",
            })
    ws, tabla = obtener_tabla(wb_valores, "TABLA CRUCE MAESTRA MP", "TablaCruceMaestraMP")
    for fila in leer_tabla(ws, tabla):
        if _texto(fila.get("Materia Prima")):
            productos.append({
                "nombre": fila["Materia Prima"],
                "codigo": _texto(fila.get("Codigo")),
                "visible": _texto(fila.get("Nombre Visible")),
                "movimiento": "compra",
                "origen": "TABLA CRUCE MAESTRA MP",
            })
    return productos


def _parecido(a, b):
    return SequenceMatcher(None, a.casefold(), b.casefold()).ratio()


def buscar_opciones(productos, texto, maximo=8):
    """Autocompletado para el SaaS: devuelve las opciones que mejor coinciden con lo escrito.

    Busca en el nombre tecnico, el nombre corto ('Harina', 'leche') y el codigo,
    sin importar mayusculas. Primero lo que contiene el texto; despues lo parecido.
    """
    buscado = _texto(texto).casefold()
    if not buscado:
        return [(p["codigo"], _texto(p["nombre"])) for p in productos][:maximo]

    def puntaje(p):
        campos = [_texto(p["nombre"]).casefold(), p.get("visible", "").casefold(), p["codigo"].casefold()]
        if any(c.startswith(buscado) for c in campos if c):
            return 3.0
        if any(buscado in c for c in campos if c):
            return 2.0
        return max(_parecido(buscado, c) for c in campos if c)

    ranking = sorted(productos, key=puntaje, reverse=True)
    return [(p["codigo"], _texto(p["nombre"])) for p in ranking if puntaje(p) >= 0.5][:maximo]


def buscar_productos_inventario(texto, maximo=8):
    """Version lista para usar desde el SaaS: lee el Excel y devuelve [(codigo, nombre), ...]."""
    from conexion_drive import leer_excel
    return buscar_opciones(leer_productos_inventario(leer_excel()), texto, maximo)


def resolver_producto_inventario(productos, escrito, max_sugerencias=3):
    """Busca por codigo o por nombre EXACTO. Si no hay match, lanza error con sugerencias."""
    buscado = _texto(escrito)
    for item in productos:
        if (item["codigo"] and item["codigo"].casefold() == buscado.casefold()) or _texto(item["nombre"]) == buscado:
            return item
    ranking = sorted(productos, key=lambda p: _parecido(buscado, _texto(p["nombre"])), reverse=True)
    sugerencias = [
        f"{_texto(p['nombre'])} ({p['codigo']})" for p in ranking[:max_sugerencias]
        if _parecido(buscado, _texto(p["nombre"])) >= 0.6
    ]
    raise ProductoNoReconocido(buscado, sugerencias)


# ----------------------------------------------------------------------------
# Escritura
# ----------------------------------------------------------------------------
def _como_fecha(valor):
    if valor is None:
        hoy = ahora_local()
        return datetime(hoy.year, hoy.month, hoy.day)
    if isinstance(valor, datetime):
        return datetime(valor.year, valor.month, valor.day)
    if isinstance(valor, date):
        return datetime(valor.year, valor.month, valor.day)
    return datetime.fromisoformat(str(valor))          # acepta '2026-09-23'


def _ya_existe(registros, fecha, factura, nombre, cantidad):
    for fila in registros:
        misma_fecha = isinstance(fila.get("Fecha"), datetime) and fila["Fecha"].date() == fecha.date()
        mismo_producto = (fila.get("Compra externa") == nombre and fila.get("Cantidad Compras") == cantidad) or \
                         (fila.get("Producción propia") == nombre and fila.get("Cantidad Producción") == cantidad)
        if misma_fecha and _texto(fila.get("Factura")) == _texto(factura) and mismo_producto:
            return fila["_fila"]
    return None


def escribir_fila_ingreso(ws, tabla, fecha, factura, nombre, cantidad, es_produccion, proveedor=None):
    """Agrega UNA fila a TablaIngresos con sus formulas. Devuelve el numero de fila.

    La usan registrar_ingresos() y tambien el Grupo 2 (hornadas), que puede
    escribir cantidades NEGATIVAS como correccion al anular una hornada.
    """
    valores = {"Fecha": fecha, "Factura": factura,
               "Proveedor": "Interno" if es_produccion else proveedor}
    if es_produccion:
        valores.update({"Producción propia": nombre, "Cantidad Producción": cantidad})
    else:
        valores.update({"Compra externa": nombre, "Cantidad Compras": cantidad})

    fila = agregar_fila(ws, tabla, valores, columnas_clave=["Fecha"])

    # Formulas de la fila (texto identico al de las filas existentes).
    # Solo se escribe el Total del lado que se uso: si no, quedaria
    # vacio x vacio = #VALUE! en el lado sin datos.
    ws[f"E{fila}"] = FORMULA_PRECIO_COMPRA.format(r=fila)
    ws[f"I{fila}"] = FORMULA_PRECIO_PRODUCCION.format(r=fila)
    ws[f"L{fila}"] = FORMULA_CATEGORIA.format(r=fila)
    if es_produccion:
        ws[f"J{fila}"] = FORMULA_TOTAL_PRODUCCION
    else:
        ws[f"F{fila}"] = FORMULA_TOTAL_COMPRA.format(r=fila)
    return fila


def registrar_ingresos(lineas, factura, proveedor=None, fecha=None):
    """Registra VARIAS lineas de un mismo documento (factura o lote) en un solo guardado.

    lineas:    [{'producto': 'Ha01', 'cantidad': 10}, {'producto': 'P01', 'cantidad': 500}]
               (codigo o nombre exacto)
    factura:   codigo de factura (ej. 'F004') o de lote de produccion (ej. 'L004')
    proveedor: nombre del proveedor; para produccion propia se usa 'Interno'
    fecha:     fecha del movimiento (date, datetime o '2026-09-23'); por defecto, hoy

    Si UNA linea tiene un nombre no reconocido, no se escribe NINGUNA
    (la factura se registra completa o no se registra).
    """
    if not lineas:
        raise ValueError("No hay lineas para registrar.")
    for linea in lineas:
        cantidad = linea.get("cantidad")
        if isinstance(cantidad, bool) or not isinstance(cantidad, (int, float)) or cantidad <= 0:
            raise ValueError(f"Cantidad invalida para '{linea.get('producto')}': {cantidad!r}")
    fecha = _como_fecha(fecha)

    def _editar(wb, wb_valores):
        productos = leer_productos_inventario(wb_valores)
        resueltas = [(resolver_producto_inventario(productos, l["producto"]), l["cantidad"]) for l in lineas]

        ws, tabla = obtener_tabla(wb, HOJA, TABLA)
        existentes = leer_tabla(ws, tabla)
        resultado = {"registradas": [], "duplicadas": []}

        for item, cantidad in resueltas:
            nombre = item["nombre"]
            repetida = _ya_existe(existentes, fecha, factura, nombre, cantidad)
            if repetida:
                resultado["duplicadas"].append({"producto": nombre, "fila": repetida})
                continue

            es_produccion = item["movimiento"] == "produccion"
            fila = escribir_fila_ingreso(ws, tabla, fecha, factura, nombre, cantidad, es_produccion, proveedor)

            existentes.append({"_fila": fila, "Fecha": fecha, "Factura": factura,
                               "Compra externa": None if es_produccion else nombre,
                               "Cantidad Compras": None if es_produccion else cantidad,
                               "Producción propia": nombre if es_produccion else None,
                               "Cantidad Producción": cantidad if es_produccion else None})
            resultado["registradas"].append({"producto": nombre, "cantidad": cantidad, "fila": fila,
                                             "movimiento": item["movimiento"], "origen": item["origen"]})

        if not resultado["registradas"]:
            return {"hubo_cambios": False, "estado": "DUPLICADO", **resultado}

        registrar_evento(
            wb, modulo="Ingreso de Inventario", tipo_evento="INGRESO_INVENTARIO", entidad="TablaIngresos",
            id_entidad=factura, accion="INSERT", estado="REGISTRADO",
            detalle="; ".join(f"{r['cantidad']} x {r['producto']} ({r['movimiento']}, fila {r['fila']})"
                              for r in resultado["registradas"]),
        )
        return {"hubo_cambios": True, "estado": "REGISTRADO", **resultado}

    return editar_excel(_editar, descripcion=f"Ingreso de inventario {factura}")


def registrar_ingreso(producto, cantidad, factura, proveedor=None, fecha=None):
    """Atajo para una sola linea."""
    return registrar_ingresos([{"producto": producto, "cantidad": cantidad}], factura, proveedor, fecha)
