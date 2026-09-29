"""Catalogo de productos y combos, leido del Excel maestro (solo lectura).

Fuente unica de nombres y precios para TODO el Grupo 1 (y para el bot):
  - TABLA CRUCE MAESTRA  (tabla TablaProductos): productos propios y de terceros
  - Tabla Combos Dinamicos (A4:L...):            combos, su precio de oferta y si estan ACTIVOS

Regla del mapa 'Ventas despacho' col F: el precio se busca en 2 pasos,
primero en combos y despues en TABLA CRUCE MAESTRA. TABLA CRUCE MAESTRA MP
(materia prima) NUNCA da precio de venta.

El producto se puede pedir por CODIGO (P01, co02...) o por el nombre EXACTO.

Reglas de ofertas confirmadas por el usuario (23-sep-2026), iguales para
despacho y venta en linea:
  - Combos: solo existen cuando su detonador los activa (ej. temperatura),
    y ya traen su descuento dentro del 'Precio de Oferta'.
  - Dias de pago (15 y 30) y feriados: 5% de descuento en TODOS los
    productos, y ese dia NO hay combos.
Las banderas 'Feriado Activo Hoy' / 'Dia de Pago Activo Hoy' se leen de
NODO INTEGRAL (Regla #4: NODO INTEGRAL es quien las calcula).
"""
from conexion_drive.tablas import leer_tabla, obtener_tabla

HOJA_COMBOS = "Tabla Combos Dinamicos"


class ProductoNoEncontrado(Exception):
    pass


class ComboInactivo(Exception):
    pass


class PrecioSinCalcular(Exception):
    pass


HOJA_INTERFACE = "Interface digital"


def _buscar_valor_por_etiqueta(ws, etiqueta, col_etiqueta, col_valor):
    for fila in range(1, ws.max_row + 1):
        texto = ws.cell(fila, col_etiqueta).value
        if isinstance(texto, str) and texto.strip().lower() == etiqueta.lower():
            return ws.cell(fila, col_valor).value
    raise KeyError(f"No encontre '{etiqueta}' en la hoja '{ws.title}'")


def _es_si(valor):
    """Solo 'SI' (o VERDADERO) cuenta como activo. '[SI/NO]' u otro texto = no activo."""
    if isinstance(valor, bool):
        return valor
    return isinstance(valor, str) and valor.strip().upper() in ("SI", "SÍ", "TRUE", "VERDADERO")


def leer_reglas_del_dia(wb_valores):
    """Banderas del dia (NODO INTEGRAL), el % de descuento general (Interface digital!B5)
    y el numero SINPE Movil de la panaderia (NODO INTEGRAL, etiqueta 'Número SINPE Móvil', m013)."""
    nodo = wb_valores["NODO INTEGRAL"]
    tarifa = _buscar_valor_por_etiqueta(nodo, "Tarifa Base KM", 2, 3)
    pct = _buscar_valor_por_etiqueta(wb_valores[HOJA_INTERFACE], "Días de Pago o feriado", 1, 2)
    feriado = _es_si(_buscar_valor_por_etiqueta(nodo, "Feriado Activo Hoy", 2, 3))
    dia_pago = _es_si(_buscar_valor_por_etiqueta(nodo, "Día de Pago Activo Hoy", 2, 3))
    try:                                   # m013 (29-sep-2026): NODO INTEGRAL fila 22, por etiqueta
        sinpe = _texto(_buscar_valor_por_etiqueta(nodo, "Número SINPE Móvil", 2, 3)) or None
    except KeyError:                       # copia sin la m013: el bot no muestra numero SINPE
        sinpe = None
    return {
        "feriado": feriado,
        "dia_pago": dia_pago,
        "descuento_general": feriado or dia_pago,          # ese dia: 5% a todo y sin combos
        "detonador_general": "FERIADO" if feriado else ("DIA_PAGO" if dia_pago else None),
        "pct_general": float(pct) if isinstance(pct, (int, float)) else 0.0,
        "tarifa_km": float(tarifa) if isinstance(tarifa, (int, float)) else 0.0,
        "sinpe_movil": sinpe,                               # numero SINPE de la panaderia (m013)
    }


def _texto(valor):
    return "" if valor is None else str(valor).strip()


def _es_numero(valor):
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


class Catalogo:
    """Se construye con el libro de VALORES (wb_valores), nunca con el de formulas."""

    def __init__(self, wb_valores):
        self.productos = self._leer_productos(wb_valores)
        self.combos = self._leer_combos(wb_valores)
        self.reglas = leer_reglas_del_dia(wb_valores)

    # ------------------------------------------------------------------ lectura
    @staticmethod
    def _leer_productos(wb_valores):
        ws, tabla = obtener_tabla(wb_valores, "TABLA CRUCE MAESTRA", "TablaProductos")
        productos = []
        for fila in leer_tabla(ws, tabla):
            if not _texto(fila.get("Nombre tecnico completo")):
                continue
            productos.append({
                "origen": "TABLA CRUCE MAESTRA",
                "nombre": fila["Nombre tecnico completo"],   # tal cual, sin recortar espacios
                "codigo": _texto(fila.get("Codigo")),
                "precio": fila.get("Precio"),
                "tipo": _texto(fila.get("Tipo")),
                "activo": True,
            })
        return productos

    @staticmethod
    def _leer_combos(wb_valores):
        ws = wb_valores[HOJA_COMBOS]
        # Ubicar la fila de encabezados buscando el texto, no por numero fijo
        fila_encabezado, columnas = None, {}
        for fila in ws.iter_rows(min_row=1, max_row=15):
            textos = {_texto(c.value): c.column for c in fila if c.value is not None}
            if "Nombre del Combo (Key)" in textos:
                fila_encabezado, columnas = fila[0].row, textos
                break
        if fila_encabezado is None:
            raise KeyError("No encontre el encabezado 'Nombre del Combo (Key)' en Tabla Combos Dinamicos")

        def valor(fila, encabezado):
            columna = columnas.get(encabezado)
            return ws.cell(fila, columna).value if columna else None

        combos = []
        for fila in range(fila_encabezado + 1, ws.max_row + 1):
            nombre = _texto(valor(fila, "Nombre del Combo (Key)"))
            codigo = _texto(valor(fila, "codigo"))
            if not nombre or not codigo:      # filas vacias o ideas sin codigo (ej. Promocion Musmani)
                continue
            combos.append({
                "origen": "Tabla Combos Dinamicos",
                "nombre": nombre,
                "codigo": codigo,
                "precio": valor(fila, "Precio de Oferta (Fórmula)"),
                "activo": _texto(valor(fila, "Detonador")).upper() == "ACTIVO",
                "tipo_detonador": _texto(valor(fila, "Tipo detonador")),
                "descuento_pct": valor(fila, "Descuento %"),
                "detalle": _texto(valor(fila, "Detalle del Combo (Descripción para el Robot)")),
            })
        return combos

    # ------------------------------------------------------------------ busqueda
    def resolver(self, producto_o_codigo, permitir_combo_inactivo=False):
        """Devuelve el diccionario del producto/combo. Lanza error si no existe."""
        buscado = _texto(producto_o_codigo)
        if not buscado:
            raise ProductoNoEncontrado("No se indico ningun producto.")

        # Paso 1: combos (por codigo sin importar mayusculas, o por nombre exacto)
        # Paso 2: TABLA CRUCE MAESTRA
        for lista in (self.combos, self.productos):
            for item in lista:
                if item["codigo"].casefold() == buscado.casefold() or _texto(item["nombre"]) == buscado:
                    return self._validar(item, permitir_combo_inactivo, self.reglas)

        raise ProductoNoEncontrado(
            f"'{buscado}' no existe en TABLA CRUCE MAESTRA ni en Tabla Combos Dinamicos. "
            "No se escribio nada (evita precios en 0 por nombres mal escritos)."
        )

    @staticmethod
    def _validar(item, permitir_combo_inactivo, reglas):
        es_combo = item["origen"] == "Tabla Combos Dinamicos"
        if es_combo and not permitir_combo_inactivo:
            if reglas["descuento_general"]:
                motivo = "feriado" if reglas["feriado"] else "dia de pago"
                raise ComboInactivo(
                    f"Hoy es {motivo}: no hay combos, aplica el descuento general a todos los productos."
                )
            if not item["activo"]:
                raise ComboInactivo(f"El combo '{item['nombre']}' ({item['codigo']}) no esta ACTIVO hoy.")
        if not _es_numero(item["precio"]):
            raise PrecioSinCalcular(
                f"El precio de '{item['nombre']}' no esta calculado en el Excel (vale {item['precio']!r}). "
                "El archivo debe recalcularse con LibreOffice antes de usarse."
            )
        return item

    def lista_para_mostrar(self):
        """Lista (codigo, nombre, precio) de lo que se puede vender hoy: productos + combos activos."""
        combos_hoy = [] if self.reglas["descuento_general"] else [c for c in self.combos if c["activo"]]
        vendibles = combos_hoy + self.productos
        return [(i["codigo"], i["nombre"], i["precio"]) for i in vendibles]
