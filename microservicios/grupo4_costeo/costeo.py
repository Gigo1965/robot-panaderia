"""Microservicio: Costeo (Grupo 4) - datos de configuracion que cambia la administracion.

Plan aprobado por el usuario el 25-sep-2026. El SaaS (o el bot) llama a estas funciones y el
robot escribe en el Excel maestro CON validacion y CON rastro:

  actualizar_precio_referencia(producto, precio)   -> 'costo pan baguette'!G28 / 'costo cupcakes'!G27
  actualizar_plaza(plaza, personas, sueldo)        -> 'planilla' B y C de las filas fijas 7 a 12
  agregar_plaza(nombre, personas, sueldo)          -> 'planilla' siguiente fila libre (13 a 32)
  actualizar_gasto(rubro, mes, monto)              -> 'gasto operacional' (solo rubros manuales)
  agregar_activo(nombre, cantidad, precio, vida)   -> 'valor activos ' filas de reserva (14 a 33)
  actualizar_activo(nombre, cantidad, precio, vida)-> 'valor activos ' (dar de baja = cantidad 0)

Reglas (decisiones del usuario):
  * Estas hojas son CONFIGURACION: el robot actualiza el valor (UPDATE), pero cada cambio deja
    un evento en RegistroEventos con el VALOR ANTERIOR -> VALOR NUEVO (no se pierde informacion).
  * Solo se escriben celdas de ENTRADA. Nunca una formula, nunca un rango de tabla, nunca se
    insertan filas (Regla #7): las filas nuevas ya existen (planilla 13-32; activos 14-33, m005).
  * planilla: los nombres de las plazas fijas (columna A, filas 7-12) NO se cambian. Solo se
    actualizan personas y sueldo. Plaza nueva = abajo, solo si el caso lo requiere.
  * Precio de referencia por debajo del costo unitario actual: se BLOQUEA (producir mas debe ser
    mas ganancia, nunca vender por debajo del costo).
  * Patente y permiso sanitario (pago unico al año): si ya hay monto en otro mes se AVISA pero
    se permite (puede ser un cambio real de fecha de pago).
  * Si el valor nuevo es igual al actual, no se sube nada (estado SIN_CAMBIO).

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo4_costeo.costeo --local precio P01 850
    py -m microservicios.grupo4_costeo.costeo --local plaza "Panadero" --sueldo 360000
    py -m microservicios.grupo4_costeo.costeo --local nueva-plaza "Repartidor" 1 300000
    py -m microservicios.grupo4_costeo.costeo --local gasto "Electricidad" 3 210000
    py -m microservicios.grupo4_costeo.costeo --local nuevo-activo "Refrigerador" 1 950000 8
    py -m microservicios.grupo4_costeo.costeo --local activo "Espátula" --cantidad 0

|| 5-oct-2026 (SaaS Grupo 4, Decisiones Pendientes #78): las seis funciones reciben el dato
OPCIONAL origen. Si no se entrega, el evento dice "Robot Python" como siempre; el SaaS entrega
"SaaS (administracion)". No cambia ninguna validacion ni ninguna celda.
Funcion nueva de LECTURA para la pantalla del SaaS: leer_costeo(wb) devuelve las dos hojas de
costo tal como las calculo el Excel (elementos del costo, costo unitario, precio de referencia
y utilidad). No escribe nada.
"""
import os
import sys

# --local: igual que las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive.drive_excel import editar_excel      # noqa: E402
from conexion_drive.tablas import registrar_evento        # noqa: E402

MODULO = "Costeo"

# ----------------------------------------------------------------------------
# Mapa de celdas (verificado en el Excel maestro despues de la m005, 25-sep-2026)
# ----------------------------------------------------------------------------
PRECIO_REFERENCIA = {
    "P01": {"hoja": "costo pan baguette", "precio": "G28", "costo": "D38", "nombre": "Baguette",
            # || 5-oct-2026: celdas de solo LECTURA para leer_costeo (formulas del Excel)
            "calculado": "G29", "utilidad_calculada": "G31", "porcentaje_calculado": "G32",
            "utilidad": "H31", "porcentaje": "H32"},
    "Pa01": {"hoja": "costo cupcakes", "precio": "G27", "costo": "D37", "nombre": "CupCake",
             "calculado": "G28", "utilidad_calculada": "G30", "porcentaje_calculado": "G31",
             "utilidad": "H30", "porcentaje": "H31"},
}
ORIGEN_ROBOT = "Robot Python"              # origen del evento cuando nadie entrega otro (#78)
# leer_costeo: la tabla "Elementos del costo" empieza debajo de la fila 9 en las dos hojas.
FILA_ELEMENTOS = 9
SECCIONES_COSTO = ("costo materiales directos", "costo mano de obra directa", "costo indirecto de fabricacion")
SINONIMOS_PRODUCTO = {"p01": "P01", "baguette": "P01", "pan baguette": "P01",
                      "pa01": "Pa01", "cupcake": "Pa01", "cupcakes": "Pa01"}

HOJA_PLANILLA = "planilla"
# Plazas fijas: nombre como lo escribe el usuario -> fila. Las dos filas "Otros " tienen el
# mismo nombre en el Excel, por eso se piden como "Otros 1" (fila 11) y "Otros 2" (fila 12).
PLAZAS_FIJAS = {"administrador": 7, "panadero": 8, "ayudante de panadero": 9, "ayudante": 9,
                "despacho": 10, "otros 1": 11, "otros 2": 12}
FILAS_PLAZA_NUEVA = range(13, 33)          # filas 13 a 32 de Tabla11 (formulas ya precargadas)

HOJA_GASTOS = "gasto operacional"
# Solo rubros MANUALES. Las filas 9, 10, 13 y 16 son formulas que vienen de 'planilla' y la
# fila 19 es el total: el robot NUNCA las escribe.
RUBROS_GASTO = {"inversion digital": 6, "inversión digital": 6, "servicio de agua": 7, "agua": 7,
                "electricidad": 8, "servicio contable": 11, "patente": 12, "alquiler local": 14,
                "alquiler": 14, "permiso sanitario": 15, "insumos higiene colaboradores": 17,
                "insumos higiene": 17, "internet": 18}
RUBROS_PAGO_UNICO = {12, 15}               # patente y permiso sanitario
COLUMNAS_MES = "BCDEFGHIJKLM"              # enero = B ... diciembre = M
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]

HOJA_ACTIVOS = "valor activos "            # OJO: el nombre real lleva un espacio al final
FILAS_ACTIVOS = range(3, 34)               # Tabla12 despues de la m005: filas 3 a 33
FILAS_RESERVA = range(14, 34)              # filas 14 a 33 (reserva creada por la m005)


class CosteoInvalido(Exception):
    """Dato rechazado por una validacion: no se escribio nada."""


# ----------------------------------------------------------------------------
# Validaciones
# ----------------------------------------------------------------------------
def _monto(valor, nombre, minimo=0.0, estricto=False):
    """Numero con 2 decimales, >= minimo (o > minimo si estricto)."""
    if isinstance(valor, bool):
        raise CosteoInvalido(f"{nombre}: {valor!r} no es un numero.")
    try:
        numero = round(float(str(valor).replace(",", "").strip()), 2)
    except (TypeError, ValueError):
        raise CosteoInvalido(f"{nombre}: {valor!r} no es un numero.") from None
    if (estricto and numero <= minimo) or (not estricto and numero < minimo):
        raise CosteoInvalido(f"{nombre} debe ser {'mayor que' if estricto else 'mayor o igual a'} "
                             f"{minimo:g} (llego {valor!r}).")
    return int(numero) if numero == int(numero) else numero


def _entero(valor, nombre, minimo=0):
    if isinstance(valor, bool):
        raise CosteoInvalido(f"{nombre}: {valor!r} no es un numero entero.")
    try:
        numero = float(str(valor).strip())
    except (TypeError, ValueError):
        raise CosteoInvalido(f"{nombre}: {valor!r} no es un numero entero.") from None
    if numero != int(numero) or numero < minimo:
        raise CosteoInvalido(f"{nombre} debe ser un entero mayor o igual a {minimo} (llego {valor!r}).")
    return int(numero)


def _texto(valor, nombre):
    texto = " ".join(str(valor or "").split())
    if not texto:
        raise CosteoInvalido(f"{nombre} no puede estar vacio.")
    if texto.startswith("="):
        raise CosteoInvalido(f"{nombre} no puede empezar con '=' (regla del proyecto).")
    return texto


def _clave(texto):
    return " ".join(str(texto or "").split()).casefold()


def _mes(valor):
    if isinstance(valor, str) and _clave(valor) in MESES:
        return MESES.index(_clave(valor)) + 1
    numero = _entero(valor, "Mes", minimo=1)
    if numero > 12:
        raise CosteoInvalido(f"Mes debe estar entre 1 y 12 (llego {valor!r}).")
    return numero


# ----------------------------------------------------------------------------
# Motor comun: escribir celdas de entrada + evento con antes -> despues
# ----------------------------------------------------------------------------
def _aplicar(wb, hoja, cambios, tipo_evento, id_entidad, accion, avisos=(), valor_referencia="",
             origen=ORIGEN_ROBOT):
    """cambios: {celda: valor_nuevo}. Escribe solo lo que cambia y registra el evento."""
    ws = wb[hoja]
    reales = []
    for celda, nuevo in cambios.items():
        anterior = ws[celda].value
        if isinstance(anterior, str) and anterior.startswith("="):
            raise CosteoInvalido(f"'{hoja}'!{celda} es una formula: el robot no la sobrescribe.")
        if anterior != nuevo:
            reales.append((celda, anterior, nuevo))
    if not reales:
        return {"hubo_cambios": False, "estado": "SIN_CAMBIO", "hoja": hoja, "cambios": [],
                "avisos": list(avisos)}
    for celda, anterior, nuevo in reales:
        ws[celda] = nuevo
    for celda, anterior, nuevo in reales:            # volver a leer y confirmar
        if ws[celda].value != nuevo:
            raise CosteoInvalido(f"Verificacion fallo en '{hoja}'!{celda}")

    detalle = " ; ".join(f"'{hoja.strip()}'!{c}: {a!r} -> {n!r}" for c, a, n in reales)
    if avisos:
        detalle += " | AVISO: " + " | ".join(avisos)
    evento = registrar_evento(wb, modulo=MODULO, tipo_evento=tipo_evento, entidad=hoja.strip(),
                              id_entidad=id_entidad, accion=accion, estado="OK",
                              valor_referencia=valor_referencia, detalle=detalle,
                              origen=origen or ORIGEN_ROBOT)
    return {"hubo_cambios": True, "estado": "OK", "hoja": hoja, "evento": evento,
            "cambios": [{"celda": c, "antes": a, "despues": n} for c, a, n in reales],
            "avisos": list(avisos)}


# ----------------------------------------------------------------------------
# 0) Lectura de las hojas de costo (pantalla del SaaS). No escribe nada.
# ----------------------------------------------------------------------------
def _numero(valor):
    return valor if isinstance(valor, (int, float)) and not isinstance(valor, bool) else None


def leer_costeo(wb):
    """Las dos hojas de costo tal como las calculo el Excel. wb = libro de VALORES (leer_excel()).

    Devuelve una lista, un producto por hoja ('costo pan baguette' y 'costo cupcakes'):
      secciones: materiales directos, mano de obra directa e indirectos, cada una con sus filas
                 (elemento, cantidad, precio, costo) y su total, tal como estan en la hoja;
      resumen:   total de costos de produccion, unidades y costo unitario;
      y las celdas del precio: el de referencia (lo digita la administracion), el calculado por
      costo x margen (solo referencia, Regla #9) y la utilidad con cada uno.
    Las filas se ubican por su TEXTO en la columna A, no por numero fijo.
    """
    productos = []
    for codigo, cfg in PRECIO_REFERENCIA.items():
        ws = wb[cfg["hoja"]]
        fila_costo = int(cfg["costo"][1:])
        secciones, resumen, actual = [], [], None
        for fila in range(FILA_ELEMENTOS + 1, fila_costo + 1):
            etiqueta = " ".join(str(ws[f"A{fila}"].value or "").split())
            if not etiqueta:
                continue
            clave = _clave(etiqueta)
            cantidad, precio, costo = (_numero(ws[f"{columna}{fila}"].value) for columna in "BCD")
            if clave in SECCIONES_COSTO:
                actual = {"titulo": etiqueta, "fila": fila, "filas": [], "total": None, "celda_total": None}
                secciones.append(actual)
            elif clave == "total" and actual is not None:
                actual["total"], actual["celda_total"] = costo, f"D{fila}"
                actual = None
            elif actual is not None:
                actual["filas"].append({"fila": fila, "elemento": etiqueta, "cantidad": cantidad,
                                        "precio": precio, "costo": costo})
            else:
                resumen.append({"fila": fila, "etiqueta": etiqueta, "valor": costo, "celda": f"D{fila}"})
        if len(secciones) != len(SECCIONES_COSTO):
            raise KeyError(f"'{cfg['hoja']}': no encontre las {len(SECCIONES_COSTO)} secciones del costo "
                           f"(encontre {len(secciones)}).")
        producto = {"codigo": codigo, "nombre": cfg["nombre"], "hoja": cfg["hoja"],
                    "secciones": secciones, "resumen": resumen,
                    "costo_unitario": _numero(ws[cfg["costo"]].value), "celda_costo": cfg["costo"],
                    "precio_referencia": _numero(ws[cfg["precio"]].value), "celda_precio": cfg["precio"]}
        for dato in ("calculado", "utilidad_calculada", "porcentaje_calculado", "utilidad", "porcentaje"):
            producto[dato] = _numero(ws[cfg[dato]].value)
            producto[f"celda_{dato}"] = cfg[dato]
        productos.append(producto)
    return productos


# ----------------------------------------------------------------------------
# 1) Precio de referencia
# ----------------------------------------------------------------------------
def actualizar_precio_referencia(producto, precio, origen=ORIGEN_ROBOT):
    """Precio de referencia de P01 (Baguette) o Pa01 (CupCake). Bloquea si queda bajo el costo.

    origen (opcional, #78): quien pide el cambio; va al evento. El SaaS entrega "SaaS (administracion)".
    """
    codigo = SINONIMOS_PRODUCTO.get(_clave(producto))
    if not codigo:
        raise CosteoInvalido(f"'{producto}' no es un producto de produccion propia (use P01 o Pa01).")
    cfg = PRECIO_REFERENCIA[codigo]
    precio = _monto(precio, "Precio de referencia", estricto=True)

    def _editar(wb, wb_valores):
        costo = wb_valores[cfg["hoja"]][cfg["costo"]].value
        avisos = []
        if isinstance(costo, (int, float)):
            if precio < round(costo, 2):
                raise CosteoInvalido(
                    f"BLOQUEADO: el precio de referencia de {cfg['nombre']} (₡{precio:,.2f}) queda por "
                    f"debajo del costo unitario actual (₡{costo:,.2f}). Regla del proyecto: producir "
                    "mas debe dar mas ganancia; no se vende por debajo del costo.")
        else:
            avisos.append(f"No se pudo comparar con el costo unitario ({costo!r}): no hay una hornada "
                          "valida cargada en la hoja de costo.")
        resultado = _aplicar(wb, cfg["hoja"], {cfg["precio"]: precio}, "CONFIG_PRECIO_REFERENCIA",
                             codigo, "UPDATE", avisos, valor_referencia=precio, origen=origen)
        # Datos para quien llama (la pantalla del SaaS); no cambian lo que se escribe.
        resultado.update(codigo=codigo, producto=cfg["nombre"], celda=cfg["precio"], precio=precio,
                         costo_unitario=costo if isinstance(costo, (int, float)) else None)
        return resultado

    return editar_excel(_editar, descripcion=f"Precio de referencia {codigo}")


# ----------------------------------------------------------------------------
# 2) Planilla
# ----------------------------------------------------------------------------
def _fila_plaza(ws, plaza):
    """Fila de una plaza: fijas por su nombre; nuevas (13-32) por el nombre escrito en A."""
    clave = _clave(plaza)
    if clave in PLAZAS_FIJAS:
        return PLAZAS_FIJAS[clave]
    if clave == "otros":
        raise CosteoInvalido("Hay dos plazas 'Otros': indique 'Otros 1' (fila 11) u 'Otros 2' (fila 12).")
    for fila in FILAS_PLAZA_NUEVA:
        if _clave(ws[f"A{fila}"].value) == clave and clave:
            return fila
    raise CosteoInvalido(f"No existe la plaza '{plaza}'. Plazas fijas: Administrador, Panadero, "
                         "Ayudante de panadero, Despacho, Otros 1, Otros 2 (o una plaza agregada).")


def actualizar_plaza(plaza, personas=None, sueldo=None, origen=ORIGEN_ROBOT):
    """Cambia el numero de personas y/o el sueldo de una plaza. El nombre (columna A) no se toca."""
    if personas is None and sueldo is None:
        raise CosteoInvalido("Indique personas y/o sueldo.")
    cambios_pedidos = {}
    if personas is not None:
        cambios_pedidos["B"] = _entero(personas, "Numero de personas")
    if sueldo is not None:
        cambios_pedidos["C"] = _monto(sueldo, "Sueldo")

    def _editar(wb, wb_valores):
        ws = wb[HOJA_PLANILLA]
        fila = _fila_plaza(ws, plaza)
        nombre = str(ws[f"A{fila}"].value).strip()
        cambios = {f"{col}{fila}": valor for col, valor in cambios_pedidos.items()}
        return _aplicar(wb, HOJA_PLANILLA, cambios, "CONFIG_PLANILLA", f"{nombre} (fila {fila})", "UPDATE",
                        origen=origen)

    return editar_excel(_editar, descripcion=f"Planilla {plaza}")


def agregar_plaza(nombre, personas, sueldo, origen=ORIGEN_ROBOT):
    """Plaza nueva en la siguiente fila libre (13-32). Las formulas de esa fila ya existen."""
    nombre = _texto(nombre, "Nombre de la plaza")
    personas = _entero(personas, "Numero de personas")
    sueldo = _monto(sueldo, "Sueldo")
    if _clave(nombre) in PLAZAS_FIJAS or _clave(nombre) == "otros":
        raise CosteoInvalido(f"'{nombre}' ya es una plaza fija: use actualizar_plaza.")

    def _editar(wb, wb_valores):
        ws = wb[HOJA_PLANILLA]
        libre = None
        for fila in FILAS_PLAZA_NUEVA:
            actual = ws[f"A{fila}"].value
            if _clave(actual) == _clave(nombre):
                raise CosteoInvalido(f"La plaza '{nombre}' ya existe (fila {fila}): use actualizar_plaza.")
            if libre is None and actual in (None, "") and ws[f"B{fila}"].value in (None, "") \
                    and ws[f"C{fila}"].value in (None, ""):
                libre = fila
        if libre is None:
            raise CosteoInvalido("No quedan filas libres en planilla (13 a 32). Hace falta una migracion.")
        cambios = {f"A{libre}": nombre, f"B{libre}": personas, f"C{libre}": sueldo}
        return _aplicar(wb, HOJA_PLANILLA, cambios, "CONFIG_PLANILLA", f"{nombre} (fila {libre})", "INSERT",
                        origen=origen)

    return editar_excel(_editar, descripcion=f"Plaza nueva {nombre}")


# ----------------------------------------------------------------------------
# 3) Gasto operacional
# ----------------------------------------------------------------------------
def actualizar_gasto(rubro, mes, monto, origen=ORIGEN_ROBOT):
    """Monto de un rubro manual en un mes (1-12 o nombre del mes)."""
    fila = RUBROS_GASTO.get(_clave(rubro))
    if fila is None:
        raise CosteoInvalido(
            f"'{rubro}' no es un rubro manual de gasto operacional. Manuales: Inversion digital, "
            "Servicio de agua, Electricidad, Servicio contable, Patente, Alquiler local, Permiso "
            "sanitario, Insumos higiene colaboradores, Internet. (Planilla, cargas sociales, poliza "
            "y aguinaldo se calculan solos desde 'planilla'.)")
    numero_mes = _mes(mes)
    monto = _monto(monto, "Monto")
    celda = f"{COLUMNAS_MES[numero_mes - 1]}{fila}"

    def _editar(wb, wb_valores):
        ws = wb[HOJA_GASTOS]
        nombre = str(ws[f"A{fila}"].value).strip()
        avisos = []
        if fila in RUBROS_PAGO_UNICO and monto:
            otros = [MESES[i].capitalize() for i, col in enumerate(COLUMNAS_MES)
                     if i != numero_mes - 1 and ws[f"{col}{fila}"].value not in (None, "", 0)]
            if otros:
                avisos.append(f"{nombre} es un pago unico al año y ya tiene monto en: {', '.join(otros)}. "
                              "Se guardo igual; revise si hay que dejar en 0 el otro mes.")
        return _aplicar(wb, HOJA_GASTOS, {celda: monto}, "CONFIG_GASTO_OPERACIONAL",
                        f"{nombre} / {MESES[numero_mes - 1]}", "UPDATE", avisos, valor_referencia=monto,
                        origen=origen)

    return editar_excel(_editar, descripcion=f"Gasto {rubro} mes {numero_mes}")


# ----------------------------------------------------------------------------
# 4) Valor activos
# ----------------------------------------------------------------------------
def _fila_activo(ws, nombre):
    for fila in FILAS_ACTIVOS:
        if _clave(ws[f"A{fila}"].value) == _clave(nombre) and _clave(nombre):
            return fila
    return None


def agregar_activo(nombre, cantidad, precio, vida_util, origen=ORIGEN_ROBOT):
    """Activo nuevo en la primera fila de reserva libre (14-33). Solo se escriben A, B, C y E."""
    nombre = _texto(nombre, "Nombre del activo")
    cantidad = _entero(cantidad, "Cantidad")
    precio = _monto(precio, "Precio")
    vida_util = _monto(vida_util, "Vida util (años)", estricto=True)

    def _editar(wb, wb_valores):
        ws = wb[HOJA_ACTIVOS]
        existente = _fila_activo(ws, nombre)
        if existente:
            raise CosteoInvalido(f"El activo '{nombre}' ya existe (fila {existente}): use actualizar_activo.")
        libre = next((f for f in FILAS_RESERVA if all(ws[f"{c}{f}"].value in (None, "") for c in "ABCE")), None)
        if libre is None:
            raise CosteoInvalido("No quedan filas de reserva en valor activos (14 a 33). Hace falta una migracion.")
        cambios = {f"A{libre}": nombre, f"B{libre}": cantidad, f"C{libre}": precio, f"E{libre}": vida_util}
        return _aplicar(wb, HOJA_ACTIVOS, cambios, "CONFIG_ACTIVO", f"{nombre} (fila {libre})", "INSERT",
                        valor_referencia=round(cantidad * precio, 2), origen=origen)

    return editar_excel(_editar, descripcion=f"Activo nuevo {nombre}")


def actualizar_activo(nombre, cantidad=None, precio=None, vida_util=None, origen=ORIGEN_ROBOT):
    """Cambia cantidad, precio y/o vida util. Dar de baja = cantidad 0 (la fila se conserva)."""
    if cantidad is None and precio is None and vida_util is None:
        raise CosteoInvalido("Indique cantidad, precio y/o vida util.")
    pedidos = {}
    if cantidad is not None:
        pedidos["B"] = _entero(cantidad, "Cantidad")
    if precio is not None:
        pedidos["C"] = _monto(precio, "Precio")
    if vida_util is not None:
        pedidos["E"] = _monto(vida_util, "Vida util (años)", estricto=True)

    def _editar(wb, wb_valores):
        ws = wb[HOJA_ACTIVOS]
        fila = _fila_activo(ws, nombre)
        if fila is None:
            raise CosteoInvalido(f"No existe el activo '{nombre}' en valor activos.")
        real = str(ws[f"A{fila}"].value).strip()
        accion = "BAJA (cantidad 0)" if pedidos.get("B") == 0 else "UPDATE"
        cambios = {f"{col}{fila}": valor for col, valor in pedidos.items()}
        return _aplicar(wb, HOJA_ACTIVOS, cambios, "CONFIG_ACTIVO", f"{real} (fila {fila})", accion,
                        origen=origen)

    return editar_excel(_editar, descripcion=f"Activo {nombre}")


# ----------------------------------------------------------------------------
# Uso manual desde la Terminal
# ----------------------------------------------------------------------------
def _mostrar(resultado):
    print(f"Estado: {resultado['estado']}")
    for c in resultado.get("cambios", []):
        print(f"  {c['celda']}: {c['antes']!r} -> {c['despues']!r}")
    for aviso in resultado.get("avisos", []):
        print(f"  AVISO: {aviso}")
    if resultado.get("evento"):
        print(f"  Evento: {resultado['evento']}")


def _main(argv):
    import argparse
    p = argparse.ArgumentParser(prog="costeo", description="Configuracion del Grupo 4 (costeo)")
    p.add_argument("--local", action="store_true", help="usar la copia datos_prueba")
    sub = p.add_subparsers(dest="accion", required=True)
    s = sub.add_parser("precio"); s.add_argument("producto"); s.add_argument("precio")
    s = sub.add_parser("plaza"); s.add_argument("plaza"); s.add_argument("--personas"); s.add_argument("--sueldo")
    s = sub.add_parser("nueva-plaza"); s.add_argument("nombre"); s.add_argument("personas"); s.add_argument("sueldo")
    s = sub.add_parser("gasto"); s.add_argument("rubro"); s.add_argument("mes"); s.add_argument("monto")
    s = sub.add_parser("nuevo-activo"); s.add_argument("nombre"); s.add_argument("cantidad")
    s.add_argument("precio"); s.add_argument("vida_util")
    s = sub.add_parser("activo"); s.add_argument("nombre"); s.add_argument("--cantidad")
    s.add_argument("--precio"); s.add_argument("--vida-util", dest="vida_util")
    a = p.parse_args(argv)

    from conexion_drive import config
    print(f"Excel de destino: {'COPIA LOCAL ' + str(config.EXCEL_LOCAL_PATH) if config.ORIGEN_EXCEL == 'local' else 'GOOGLE DRIVE (real)'}")
    if a.accion == "precio":
        return actualizar_precio_referencia(a.producto, a.precio)
    if a.accion == "plaza":
        return actualizar_plaza(a.plaza, a.personas, a.sueldo)
    if a.accion == "nueva-plaza":
        return agregar_plaza(a.nombre, a.personas, a.sueldo)
    if a.accion == "gasto":
        return actualizar_gasto(a.rubro, a.mes, a.monto)
    if a.accion == "nuevo-activo":
        return agregar_activo(a.nombre, a.cantidad, a.precio, a.vida_util)
    return actualizar_activo(a.nombre, a.cantidad, a.precio, a.vida_util)


if __name__ == "__main__":
    try:
        _mostrar(_main(sys.argv[1:]))
    except CosteoInvalido as error:
        print(f"NO SE CAMBIO NADA. {error}")
        sys.exit(1)
