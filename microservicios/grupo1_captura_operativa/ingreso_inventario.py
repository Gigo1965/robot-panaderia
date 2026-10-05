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

El robot NUNCA inventa ni digita precios: el precio viene SIEMPRE de la
TABLA CRUCE MAESTRA / MP. 'Precio Unitario2' (produccion), los Totales y la
Categoria son formulas que se calculan solas. Como las filas nuevas no heredan
formulas, el robot escribe el TEXTO de cada formula, igual al de las filas
existentes.

#32 (V15, 27-sep-2026) - COSTO DE LA COMPRA CONGELADO:
  Antes, 'Precio unitario' (columna E) era una formula viva a las tablas
  maestras: cuando precios_proveedores.py cambiaba un precio, TODAS las compras
  pasadas se recalculaban. Ahora, al registrar una compra, el robot toma el
  costo que la TCM tiene EN ESE MOMENTO (misma busqueda que la formula:
  primero TablaProductos 'Costo Unitario (Robot)', si no, TablaCruceMaestraMP
  'Costo Unitario (Scraping)', siempre por NOMBRE) y lo escribe como NUMERO.
  Las compras nuevas toman el precio nuevo; las ya registradas no cambian.
  Si la TCM no tiene un costo mayor que 0 para el producto, la factura
  completa se rechaza sin escribir (CostoNoDisponible). La produccion propia
  (columna I) sigue como formula (decision del usuario). Las compras viejas
  se congelaron con la migracion m012.

CAFE Ca01 / Cf01 (V15, 28-sep-2026, decision del usuario = opcion A):
  Un producto que se VENDE tal como se compra (esta en TablaProductos) se compra
  con su codigo de VENTA y la cantidad va en unidades de venta (Regla #32). Si
  el mismo nombre esta tambien en la MP con otro codigo (el cafe: Cf01 paquete en
  la TCM, Ca01 caja de 10 en la MP), el codigo de la MP se RECHAZA en compras
  (UsarCodigoDeVenta, no se escribe nada) y el mensaje indica el codigo correcto
  y cuantas unidades de venta trae la presentacion de la MP. Asi 'Ca01, 1' ya no
  se cuenta en silencio como 1 paquete. El autocompletado y las sugerencias solo
  muestran el codigo de venta. La fila de la MP NO se toca (la usa
  precios_proveedores.py). No es una regla solo para el cafe: aplica a cualquier
  producto que llegue a estar en las dos tablas.

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
# Categoria: misma formula que la m006 dejo en L4:L37 (Regla #32): lo que esta en
# TablaProductos se revende -> 'Compras Externas Netas'; si no, la MP -> 'Materia Prima'.
# (Antes de la m006 preguntaba primero por la MP y el cafe salia 'Materia Prima'.)
FORMULA_CATEGORIA = (
    '=IF(C{r}="","",IF(ISNUMBER(MATCH(C{r},TablaProductos[Nombre tecnico completo],0)),'
    '"Compras Externas Netas",IF(ISNUMBER(MATCH(C{r},TablaCruceMaestraMP[Materia Prima],0)),'
    '"Materia Prima","Compras Externas Netas")))'
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


class UsarCodigoDeVenta(ProductoNoReconocido):
    """Se uso el codigo de la MP de un producto que se vende tal como se compra (V15).

    Hereda de ProductoNoReconocido: el SaaS lo maneja igual (rechazo + sugerencia).
    """

    def __init__(self, escrito, nombre, codigo_venta, factor=None):
        self.escrito = escrito
        self.nombre = nombre
        self.codigo_venta = codigo_venta
        self.factor = factor
        self.sugerencias = [f"{nombre} ({codigo_venta})"]
        texto = (f"'{escrito}' es el codigo de TABLA CRUCE MAESTRA MP de un producto que se vende tal como "
                 f"se compra. Registre la compra con {codigo_venta} ({nombre}) y la cantidad en unidades de venta")
        if factor:
            texto += f" (1 unidad de '{escrito}' = {factor} unidades de {codigo_venta})"
        Exception.__init__(self, texto + ". No se escribio nada.")


class CostoNoDisponible(Exception):
    """La TCM no tiene un costo mayor que 0 para el producto (#32). No se escribe nada."""

    def __init__(self, nombre, costo):
        self.nombre = nombre
        self.costo = costo
        super().__init__(
            f"'{nombre}' no tiene costo valido en TABLA CRUCE MAESTRA / MP (valor: {costo!r}). "
            "No se escribio nada. Revise el precio del proveedor antes de registrar la compra."
        )


# ----------------------------------------------------------------------------
# Lista valida de productos para inventario
# ----------------------------------------------------------------------------
def _texto(valor):
    return "" if valor is None else str(valor).strip()


def _factor_presentacion(costo_mp, costo_venta):
    """Cuantas unidades de venta trae la presentacion de la MP (cafe: 20,755 / 2,075.5 = 10).

    Solo para el mensaje de rechazo; nunca convierte cantidades. None si no es un entero claro.
    """
    if not all(isinstance(c, (int, float)) and not isinstance(c, bool) and c > 0 for c in (costo_mp, costo_venta)):
        return None
    razon = costo_mp / costo_venta
    return round(razon) if round(razon) > 1 and abs(razon - round(razon)) < 0.01 else None


def leer_productos_inventario(wb_valores):
    """Todos los nombres validos para Ingreso de Inventario, con su tipo de movimiento.

    V15: si un nombre de la MP tambien esta en TablaProductos (se vende tal como se
    compra, p. ej. el cafe), la entrada de la MP lleva 'codigo_venta' (y 'factor'):
    se rechaza en compras y no se muestra en el autocompletado ni en sugerencias.
    """
    productos, venta = [], {}
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
            venta.setdefault(_texto(fila["Nombre tecnico completo"]),
                             (_texto(fila.get("Codigo")), fila.get("Costo Unitario (Robot)")))
    ws, tabla = obtener_tabla(wb_valores, "TABLA CRUCE MAESTRA MP", "TablaCruceMaestraMP")
    for fila in leer_tabla(ws, tabla):
        if _texto(fila.get("Materia Prima")):
            item = {
                "nombre": fila["Materia Prima"],
                "codigo": _texto(fila.get("Codigo")),
                "visible": _texto(fila.get("Nombre Visible")),
                "movimiento": "compra",
                "origen": "TABLA CRUCE MAESTRA MP",
            }
            if _texto(fila["Materia Prima"]) in venta:                     # V15: cafe Ca01
                codigo_venta, costo_venta = venta[_texto(fila["Materia Prima"])]
                item["codigo_venta"] = codigo_venta
                item["factor"] = _factor_presentacion(fila.get("Costo Unitario (Scraping)"), costo_venta)
            productos.append(item)
    return productos


def leer_costos_tcm(wb_valores):
    """Costo actual de cada producto POR NOMBRE, con la misma prioridad que la formula
    vieja de 'Precio unitario': primero TablaProductos, despues TablaCruceMaestraMP.

    Por eso el cafe (mismo nombre en las dos tablas) toma el costo POR PAQUETE de
    TablaProductos (2,075.50), igual que antes, aunque se registre con el codigo Ca01.
    || V15 (28-sep-2026): Ca01 ya no se acepta en compras (UsarCodigoDeVenta); el cafe
    se registra con Cf01 y la cantidad en paquetes.
    """
    costos = {}
    for hoja, nombre_tabla, col_nombre, col_costo in (
        ("TABLA CRUCE MAESTRA", "TablaProductos", "Nombre tecnico completo", "Costo Unitario (Robot)"),
        ("TABLA CRUCE MAESTRA MP", "TablaCruceMaestraMP", "Materia Prima", "Costo Unitario (Scraping)"),
    ):
        ws, tabla = obtener_tabla(wb_valores, hoja, nombre_tabla)
        for fila in leer_tabla(ws, tabla):
            nombre = _texto(fila.get(col_nombre))
            if nombre and nombre not in costos:
                costos[nombre] = fila.get(col_costo)
    return costos


def _costo_congelado(nombre, costos):
    """Devuelve el costo de la TCM para congelarlo; si no es un numero mayor que 0, rechaza."""
    costo = costos.get(_texto(nombre))
    if isinstance(costo, bool) or not isinstance(costo, (int, float)) or costo <= 0:
        raise CostoNoDisponible(_texto(nombre), costo)
    return costo


def _parecido(a, b):
    return SequenceMatcher(None, a.casefold(), b.casefold()).ratio()


def buscar_opciones(productos, texto, maximo=8):
    """Autocompletado para el SaaS: devuelve las opciones que mejor coinciden con lo escrito.

    Busca en el nombre tecnico, el nombre corto ('Harina', 'leche') y el codigo,
    sin importar mayusculas. Primero lo que contiene el texto; despues lo parecido.
    """
    productos = [p for p in productos if not p.get("codigo_venta")]      # V15: sin Ca01
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
            if item.get("codigo_venta"):                                   # V15: cafe Ca01 -> use Cf01
                raise UsarCodigoDeVenta(buscado, _texto(item["nombre"]), item["codigo_venta"], item.get("factor"))
            return item
    productos = [p for p in productos if not p.get("codigo_venta")]      # sugerencias sin Ca01
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


def escribir_fila_ingreso(ws, tabla, fecha, factura, nombre, cantidad, es_produccion, proveedor=None,
                          costo_unitario=None):
    """Agrega UNA fila a TablaIngresos con sus formulas. Devuelve el numero de fila.

    La usan registrar_ingresos() y tambien el Grupo 2 (hornadas), que puede
    escribir cantidades NEGATIVAS como correccion al anular una hornada.

    costo_unitario (#32): costo de la TCM al registrar una COMPRA; se escribe en E
    como NUMERO (queda congelado). registrar_ingresos() siempre lo envia. Si una
    llamada vieja no lo trae (m009, ya aplicada), E queda con la formula de antes
    y la m012 la congela. En produccion no se usa (E queda vacia por la formula).
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
    if not es_produccion and costo_unitario is not None:
        ws[f"E{fila}"] = costo_unitario                     # #32: costo congelado (numero)
    else:
        ws[f"E{fila}"] = FORMULA_PRECIO_COMPRA.format(r=fila)
    ws[f"I{fila}"] = FORMULA_PRECIO_PRODUCCION.format(r=fila)
    ws[f"L{fila}"] = FORMULA_CATEGORIA.format(r=fila)
    if es_produccion:
        ws[f"J{fila}"] = FORMULA_TOTAL_PRODUCCION
    else:
        ws[f"F{fila}"] = FORMULA_TOTAL_COMPRA.format(r=fila)
    return fila


def registrar_ingresos(lineas, factura, proveedor=None, fecha=None, origen="Robot Python"):
    """Registra VARIAS lineas de un mismo documento (factura o lote) en un solo guardado.

    lineas:    [{'producto': 'Ha01', 'cantidad': 10}, {'producto': 'P01', 'cantidad': 500}]
               (codigo o nombre exacto)
    factura:   codigo de factura (ej. 'F004') o de lote de produccion (ej. 'L004')
    proveedor: nombre del proveedor; para produccion propia se usa 'Interno'
    fecha:     fecha del movimiento (date, datetime o '2026-09-23'); por defecto, hoy

    Si UNA linea tiene un nombre no reconocido, un codigo de la MP de un producto que
    se vende tal como se compra (V15, cafe Ca01), o una compra sin costo valido en
    la TCM (#32), no se escribe NINGUNA (la factura se registra completa o no se registra).
    El costo de cada compra se toma de la TCM en este momento y se congela.

    || 4-oct-2026 (decision del usuario, opcion C): origen es un dato OPCIONAL que solo se
    anota en la columna ORIGEN del evento (p. ej. "SaaS (administracion)"). Si no se envia,
    queda "Robot Python" como antes (EVT-000043 de la copia DEMO decia "Robot Python").
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

        # #32: costo de cada COMPRA, leido de la TCM ahora. Se revisan TODAS las lineas
        # antes de escribir: si una no tiene costo valido, no se escribe ninguna.
        costos = leer_costos_tcm(wb_valores)
        costo_de = {}
        for item, _ in resueltas:
            if item["movimiento"] != "produccion":
                costo_de[_texto(item["nombre"])] = _costo_congelado(item["nombre"], costos)

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
            costo = None if es_produccion else costo_de[_texto(nombre)]
            fila = escribir_fila_ingreso(ws, tabla, fecha, factura, nombre, cantidad, es_produccion, proveedor,
                                         costo_unitario=costo)

            existentes.append({"_fila": fila, "Fecha": fecha, "Factura": factura,
                               "Compra externa": None if es_produccion else nombre,
                               "Cantidad Compras": None if es_produccion else cantidad,
                               "Producción propia": nombre if es_produccion else None,
                               "Cantidad Producción": cantidad if es_produccion else None})
            resultado["registradas"].append({"producto": nombre, "cantidad": cantidad, "fila": fila,
                                             "movimiento": item["movimiento"], "origen": item["origen"],
                                             "costo_unitario": costo})

        if not resultado["registradas"]:
            return {"hubo_cambios": False, "estado": "DUPLICADO", **resultado}

        registrar_evento(
            wb, modulo="Ingreso de Inventario", tipo_evento="INGRESO_INVENTARIO", entidad="TablaIngresos",
            id_entidad=factura, accion="INSERT", estado="REGISTRADO", origen=origen,
            detalle="; ".join(
                f"{r['cantidad']} x {r['producto']} ({r['movimiento']}, fila {r['fila']}"
                + (f", costo congelado {r['costo_unitario']:,.2f})" if r["costo_unitario"] is not None else ")")
                for r in resultado["registradas"]),
        )
        return {"hubo_cambios": True, "estado": "REGISTRADO", **resultado}

    return editar_excel(_editar, descripcion=f"Ingreso de inventario {factura}")


def registrar_ingreso(producto, cantidad, factura, proveedor=None, fecha=None):
    """Atajo para una sola linea."""
    return registrar_ingresos([{"producto": producto, "cantidad": cantidad}], factura, proveedor, fecha)
