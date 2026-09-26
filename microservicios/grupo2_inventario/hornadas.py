"""Microservicio: Hornadas (Grupo 2 - Inventario).

Un solo registro del panadero en el SaaS hace TODO esto en el mismo guardado
(decisiones del usuario, 23-sep-2026):

  1. Hoja de costo ('costo pan baguette' o 'costo cupcakes'): el robot escribe
     la cantidad (G4), la fecha (C7) y la hora (C5 / C6), y limpia G2.
     Asi la receta calcula el costo real de esa hornada.
  2. 'Historico Diario': agrega UNA fila nueva con los 9 valores ya calculados
     (nunca formulas), con el nombre que busca 'flujo caja anual', tomado de la
     celda A6 de la hoja de costo. Desde la migracion m008 (26-sep-2026) A6 dice el
     nombre de la TABLA CRUCE MAESTRA ('Baguette' / 'CupCake'): un solo nombre por
     producto en todo el documento (antes: 'pan baguette' / 'cupcakes').
  3. 'Ingreso de Inventario': agrega la produccion propia con el nombre de la
     tabla maestra ('Baguette' / 'CupCake'), para que sume al inventario.
  4. 'Consumo de Materia Prima' (migracion m002): una fila por cada materia prima
     que gasta la receta (harina, sal, levadura...), para rebajarla del inventario.

Anular una hornada (anular_hornada):
  - Escribe G2 = 0 en la hoja de costo (junto al texto 'ORDEN ANULADA').
  - Agrega filas de CORRECCION con cantidad NEGATIVA en 'Historico Diario', en
    'Ingreso de Inventario' y en 'Consumo de Materia Prima': la hornada original
    queda como rastro y la suma neta queda en cero (inventario y 'flujo caja anual').

Duplicados: misma Fecha + Hora + Producto + Cantidad en 'Historico Diario'.
"""
from datetime import datetime, time

from conexion_drive import editar_excel
from conexion_drive.recalculo import valores_recalculados
from conexion_drive.tablas import agregar_fila, leer_tabla, obtener_tabla, registrar_evento
from conexion_drive.tiempo import a_hora_local, ahora_local
from microservicios.grupo1_captura_operativa.ingreso_inventario import escribir_fila_ingreso

# Celdas de cada hoja de costo (verificadas en el Excel maestro v68 y en el libro de mapas).
# "precio" = Precio de referencia que digita la administracion (lo crea migraciones/m001_grupo2):
# el precio de venta NO cambia con el tamaño de la hornada; el costo si. Asi, producir mas
# se ve como mas utilidad en 'Historico Diario' y 'flujo caja anual'.
HOJAS_COSTO = {
    "P01": {"hoja": "costo pan baguette", "fecha": "C7", "hora": "C5", "cantidad": "G4",
            "anulada": "G2", "nombre": "A6", "costo": "D38", "precio": "G28"},
    "Pa01": {"hoja": "costo cupcakes", "fecha": "C7", "hora": "C6", "cantidad": "G4",
             "anulada": "G2", "nombre": "A6", "costo": "D37", "precio": "G27"},
}
# Se aceptan tambien los nombres viejos ('pan baguette' / 'cupcakes', antes de la m008).
SINONIMOS = {"p01": "P01", "baguette": "P01", "pan baguette": "P01",
             "pa01": "Pa01", "cupcake": "Pa01", "cupcakes": "Pa01"}


HOJA_CONSUMO, TABLA_RECETA, TABLA_CONSUMO = "Consumo de Materia Prima", "TablaRecetaConsumo", "TablaConsumoMP"


class HornadaInvalida(Exception):
    pass


# ----------------------------------------------------------------------------
# Ayudas
# ----------------------------------------------------------------------------
def _codigo(producto):
    codigo = SINONIMOS.get(str(producto or "").strip().casefold())
    if not codigo:
        raise HornadaInvalida(f"'{producto}' no es un producto de produccion propia (use P01 o Pa01).")
    return codigo


def _nombre_maestro(wb_valores, codigo):
    ws, tabla = obtener_tabla(wb_valores, "TABLA CRUCE MAESTRA", "TablaProductos")
    for fila in leer_tabla(ws, tabla):
        if str(fila.get("Codigo") or "").strip() == codigo:
            return fila["Nombre tecnico completo"]
    raise HornadaInvalida(f"El codigo {codigo} no esta en TABLA CRUCE MAESTRA.")


def _como_hora(valor):
    if isinstance(valor, datetime):
        return valor.time().replace(second=0, microsecond=0)
    if isinstance(valor, time):
        return valor.replace(second=0, microsecond=0)
    return None


def _buscar_en_historico(ws, tabla, fecha, hora, nombre, cantidad):
    for fila in leer_tabla(ws, tabla):
        f = fila.get("Fecha")
        if (isinstance(f, datetime) and f.date() == fecha.date() and _como_hora(fila.get("Hora")) == hora
                and fila.get("Producto") == nombre and fila.get("Cantidad producida") == cantidad):
            return fila
    return None


def _escribir_consumo(wb, wb_calculado, fecha, lote, codigo, nombre_inv, cantidad):
    """Una fila por materia prima de la receta. Devuelve la lista de filas escritas."""
    if HOJA_CONSUMO not in wb.sheetnames:
        raise HornadaInvalida("Falta la hoja 'Consumo de Materia Prima': aplique la migracion m002_consumo_mp.")
    ws_r, tabla_r = obtener_tabla(wb_calculado, HOJA_CONSUMO, TABLA_RECETA)
    ws_c, tabla_c = obtener_tabla(wb, HOJA_CONSUMO, TABLA_CONSUMO)
    filas = []
    for linea in leer_tabla(ws_r, tabla_r):
        if str(linea.get("Codigo") or "").strip() != codigo:
            continue
        por_unidad, contenido = linea.get("Cantidad por unidad"), linea.get("Contenido por paquete")
        if not isinstance(por_unidad, (int, float)) or not isinstance(contenido, (int, float)) or contenido <= 0:
            raise HornadaInvalida(f"Receta incompleta para {linea.get('Materia prima')}: {por_unidad!r} / {contenido!r}")
        total = round(cantidad * por_unidad, 4)
        filas.append(agregar_fila(ws_c, tabla_c, {
            "Fecha": fecha, "Lote": lote, "Producto horneado": nombre_inv, "Cantidad horneada": cantidad,
            "Materia prima": linea["Materia prima"], "Consumo por unidad": por_unidad, "Consumo total": total,
            "Contenido por paquete": contenido, "Paquetes consumidos": round(total / contenido, 4),
        }, columnas_clave=["Fecha"]))
    return filas


def _revertir_consumo(wb, fecha, lote_original, lote_correccion):
    ws_c, tabla_c = obtener_tabla(wb, HOJA_CONSUMO, TABLA_CONSUMO)
    originales = [f for f in leer_tabla(ws_c, tabla_c) if f.get("Lote") == lote_original]
    filas = []
    for f in originales:
        filas.append(agregar_fila(ws_c, tabla_c, {
            "Fecha": fecha, "Lote": lote_correccion, "Producto horneado": f["Producto horneado"],
            "Cantidad horneada": -f["Cantidad horneada"], "Materia prima": f["Materia prima"],
            "Consumo por unidad": f["Consumo por unidad"], "Consumo total": -f["Consumo total"],
            "Contenido por paquete": f["Contenido por paquete"], "Paquetes consumidos": -f["Paquetes consumidos"],
        }, columnas_clave=["Fecha"]))
    return filas


def _lote(fecha, hora, codigo):
    return f"H-{fecha:%Y%m%d}-{hora:%H%M}-{codigo}"


def _fila_historico(fecha, hora, nombre, cantidad, costo, precio):
    venta = round(cantidad * precio, 2)
    costo_total = round(cantidad * costo, 2)
    return {
        "Fecha": fecha, "Hora": hora, "Producto": nombre, "Cantidad producida": cantidad,
        "Costo unitario": costo, "Precio unitario": precio, "Venta total": venta,
        "Costo total producción": costo_total, "Utilidad bruta": round(venta - costo_total, 2),
    }


# ----------------------------------------------------------------------------
# Funciones publicas
# ----------------------------------------------------------------------------
def registrar_hornada(producto, cantidad, momento=None):
    """Registra una hornada. producto: 'P01'/'Baguette' o 'Pa01'/'CupCake'."""
    if isinstance(cantidad, bool) or not isinstance(cantidad, int) or cantidad <= 0:
        raise HornadaInvalida(f"La cantidad debe ser un entero mayor que 0 (llego {cantidad!r}). "
                              "Para anular una hornada use anular_hornada().")
    codigo = _codigo(producto)
    cfg = HOJAS_COSTO[codigo]
    momento = a_hora_local(momento) if momento else ahora_local()
    fecha = datetime(momento.year, momento.month, momento.day)
    hora = time(momento.hour, momento.minute)

    def _editar(wb, wb_valores):
        nombre_hist = wb_valores[cfg["hoja"]][cfg["nombre"]].value      # A6: 'Baguette' / 'CupCake' (m008)
        nombre_inv = _nombre_maestro(wb_valores, codigo)                # 'Baguette' / 'CupCake'
        ws_h, tabla_h = obtener_tabla(wb, "Historico Diario", "HistoricoHorneadas")
        if _buscar_en_historico(ws_h, tabla_h, fecha, hora, nombre_hist, cantidad):
            return {"hubo_cambios": False, "estado": "DUPLICADO", "producto": nombre_hist, "cantidad": cantidad}

        # 1) Hoja de costo: la receta recalcula con la cantidad nueva
        ws_c = wb[cfg["hoja"]]
        ws_c[cfg["cantidad"]] = cantidad
        ws_c[cfg["fecha"]] = fecha
        ws_c[cfg["hora"]] = hora
        ws_c[cfg["anulada"]] = None
        wb_calculado = valores_recalculados(wb)
        calculado = wb_calculado[cfg["hoja"]]
        costo, precio = calculado[cfg["costo"]].value, calculado[cfg["precio"]].value
        if not isinstance(costo, (int, float)) or not isinstance(precio, (int, float)):
            raise HornadaInvalida(f"La hoja '{cfg['hoja']}' no calculo costo/precio ({costo!r}, {precio!r}).")
        costo, precio = round(costo, 2), round(precio, 2)

        # 2) Historico Diario (valores fijos)
        valores = _fila_historico(fecha, hora, nombre_hist, cantidad, costo, precio)
        fila_h = agregar_fila(ws_h, tabla_h, valores, columnas_clave=["Fecha"])

        # 3) Ingreso de Inventario (produccion propia)
        ws_i, tabla_i = obtener_tabla(wb, "Ingreso de Inventario", "TablaIngresos")
        lote = _lote(fecha, hora, codigo)
        fila_i = escribir_fila_ingreso(ws_i, tabla_i, fecha, lote, nombre_inv, cantidad, es_produccion=True)

        # 4) Consumo de materia prima segun la receta
        filas_c = _escribir_consumo(wb, wb_calculado, fecha, lote, codigo, nombre_inv, cantidad)

        registrar_evento(wb, modulo="Hornadas", tipo_evento="HORNADA", entidad=cfg["hoja"], id_entidad=lote,
                         accion="INSERT", estado="REGISTRADA", valor_referencia=valores["Venta total"],
                         detalle=f"{cantidad} x {nombre_hist} a costo {costo} | Historico fila {fila_h} | "
                                 f"Ingreso fila {fila_i} | Consumo filas {filas_c}")
        return {"hubo_cambios": True, "estado": "REGISTRADA", "lote": lote, "producto": nombre_hist,
                "cantidad": cantidad, "costo_unitario": costo, "precio_unitario": precio,
                "venta_total": valores["Venta total"], "utilidad_bruta": valores["Utilidad bruta"],
                "fila_historico": fila_h, "fila_ingreso": fila_i, "filas_consumo": filas_c}

    return editar_excel(_editar, descripcion=f"Hornada {cantidad} x {codigo}")


def anular_hornada(producto):
    """Anula la hornada que esta cargada ahora en la hoja de costo (G2 = 0 + correcciones negativas)."""
    codigo = _codigo(producto)
    cfg = HOJAS_COSTO[codigo]

    def _editar(wb, wb_valores):
        hoja_v = wb_valores[cfg["hoja"]]
        if hoja_v[cfg["anulada"]].value == 0:
            return {"hubo_cambios": False, "estado": "YA_ESTABA_ANULADA"}
        cantidad = hoja_v[cfg["cantidad"]].value
        fecha_v, hora = hoja_v[cfg["fecha"]].value, _como_hora(hoja_v[cfg["hora"]].value)
        if not isinstance(cantidad, (int, float)) or cantidad <= 0 or not isinstance(fecha_v, datetime):
            raise HornadaInvalida(f"No hay una hornada cargada en '{cfg['hoja']}' para anular.")
        cantidad, fecha = int(cantidad), datetime(fecha_v.year, fecha_v.month, fecha_v.day)
        nombre_hist = hoja_v[cfg["nombre"]].value
        nombre_inv = _nombre_maestro(wb_valores, codigo)

        ws_h, tabla_h = obtener_tabla(wb, "Historico Diario", "HistoricoHorneadas")
        original = _buscar_en_historico(ws_h, tabla_h, fecha, hora, nombre_hist, cantidad)
        if original is None:
            raise HornadaInvalida(f"La hornada {cantidad} x {nombre_hist} del {fecha:%d/%m/%Y} {hora} "
                                  "no esta en Historico Diario.")

        wb[cfg["hoja"]][cfg["anulada"]] = 0
        correccion = _fila_historico(fecha, hora, nombre_hist, -cantidad,
                                     original["Costo unitario"], original["Precio unitario"])
        fila_h = agregar_fila(ws_h, tabla_h, correccion, columnas_clave=["Fecha"])

        ws_i, tabla_i = obtener_tabla(wb, "Ingreso de Inventario", "TablaIngresos")
        lote = "ANULA-" + _lote(fecha, hora, codigo)
        fila_i = escribir_fila_ingreso(ws_i, tabla_i, fecha, lote, nombre_inv, -cantidad, es_produccion=True)
        filas_c = _revertir_consumo(wb, fecha, _lote(fecha, hora, codigo), lote)

        registrar_evento(wb, modulo="Hornadas", tipo_evento="ANULACION_HORNADA", entidad=cfg["hoja"],
                         id_entidad=lote, accion="G2=0 + correccion negativa", estado="ANULADA",
                         valor_referencia=correccion["Venta total"],
                         detalle=f"-{cantidad} x {nombre_hist} | Historico fila {fila_h} | Ingreso fila {fila_i} | "
                                 f"Consumo filas {filas_c}")
        return {"hubo_cambios": True, "estado": "ANULADA", "lote": lote, "cantidad_corregida": -cantidad,
                "fila_historico": fila_h, "fila_ingreso": fila_i, "filas_consumo": filas_c}

    return editar_excel(_editar, descripcion=f"Anular hornada {codigo}")
