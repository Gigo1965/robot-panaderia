"""Microservicio: Ventas despacho (Grupo 1 - Captura operativa).

Sigue el 'Diseño del robot' de Ventas despacho del libro de mapas V5:
  1. El despachador confirma una venta en el SaaS (producto por codigo + cantidad).
  2. El robot revisa duplicados: Fecha + Hora + Producto + Cantidad.
  3. Calcula el precio (combo primero, TABLA CRUCE MAESTRA despues).
     En feriado o dia de pago (NODO INTEGRAL): no hay combos y el precio
     unitario que se escribe ya lleva el 5% de descuento general
     (misma regla que la venta en linea, confirmada 23-sep-2026).
  4. Agrega la fila nueva a Tabla6 (INSERT) con valores FIJOS:
       Fecha, Hora, Producto, Cantidad, Precio  -> valores fijos
       Total                                    -> formula de la misma fila
       Estado de orden                          -> 'ORDEN CORRECTA'
  5. Anulacion: unica excepcion al "solo INSERT"; cambia SOLO 'Estado de orden'.
Cada accion queda anotada en la hoja RegistroEventos.
"""
from datetime import datetime, time

from conexion_drive import editar_excel
from conexion_drive.tablas import agregar_fila, columnas_de_tabla, leer_tabla, obtener_tabla, registrar_evento
from conexion_drive.tiempo import a_hora_local, ahora_local

from .catalogo import Catalogo

HOJA = "Ventas despacho"
TABLA = "Tabla6"
ORDEN_CORRECTA = "ORDEN CORRECTA"
ORDEN_ANULADA = "ORDEN ANULADA"
FORMULA_TOTAL = "=Tabla6[[#This Row],[Cantidad]]*Tabla6[[#This Row],[Precio]]"


# ----------------------------------------------------------------------------
# Ayudas internas
# ----------------------------------------------------------------------------
def _validar_cantidad(cantidad):
    if isinstance(cantidad, bool) or not isinstance(cantidad, int) or cantidad <= 0:
        raise ValueError(f"La cantidad debe ser un numero entero mayor que 0 (llego {cantidad!r}).")
    return cantidad


def _separar_fecha_hora(momento):
    momento = a_hora_local(momento) if momento else ahora_local()
    fecha = datetime(momento.year, momento.month, momento.day)          # medianoche, como el historico
    hora = time(momento.hour, momento.minute, momento.second)
    return fecha, hora


def _misma_fecha(valor, fecha):
    return isinstance(valor, datetime) and valor.date() == fecha.date()


def _misma_hora(valor, hora):
    if isinstance(valor, datetime):
        valor = valor.time()
    return isinstance(valor, time) and valor.replace(microsecond=0) == hora


def _buscar_venta(ws, tabla, fecha, hora, nombre, cantidad):
    for fila in leer_tabla(ws, tabla):
        if (_misma_fecha(fila.get("Fecha"), fecha) and _misma_hora(fila.get("Hora"), hora)
                and fila.get("Producto") == nombre and fila.get("Cantidad") == cantidad):
            return fila
    return None


# ----------------------------------------------------------------------------
# Funciones publicas
# ----------------------------------------------------------------------------
def registrar_venta_despacho(producto, cantidad, momento=None):
    """Registra una venta de mostrador.

    producto: codigo (ej. 'P01', 'co02') o nombre exacto de la tabla maestra.
    cantidad: numero entero > 0.
    momento:  fecha y hora de la venta que envia el SaaS (datetime). Si trae zona
              horaria (ej. UTC) se convierte a hora de Costa Rica; si no trae zona,
              se asume que ya es hora de Costa Rica. Si no se envia, se usa la hora
              actual de Costa Rica. IMPORTANTE: el SaaS
              debe enviar siempre el momento original, para que un reintento
              se reconozca como duplicado.
    """
    cantidad = _validar_cantidad(cantidad)
    fecha, hora = _separar_fecha_hora(momento)

    def _editar(wb, wb_valores):
        catalogo = Catalogo(wb_valores)
        item = catalogo.resolver(producto)
        nombre, precio = item["nombre"], round(float(item["precio"]), 2)
        nota_descuento = ""
        if catalogo.reglas["descuento_general"]:               # combos ya fueron rechazados
            pct = catalogo.reglas["pct_general"]
            precio = round(precio * (1 - pct), 2)
            nota_descuento = f", {pct:.0%} {catalogo.reglas['detonador_general']}"
        ws, tabla = obtener_tabla(wb, HOJA, TABLA)

        existente = _buscar_venta(ws, tabla, fecha, hora, nombre, cantidad)
        if existente:
            return {"hubo_cambios": False, "estado": "DUPLICADO", "fila": existente["_fila"],
                    "producto": nombre, "cantidad": cantidad}

        fila = agregar_fila(ws, tabla, {
            "Fecha": fecha,
            "Hora": hora,
            "Producto": nombre,
            "Cantidad": cantidad,
            "Precio": precio,
            "Total": FORMULA_TOTAL,
            "Estado de orden": ORDEN_CORRECTA,
        }, columnas_clave=["Fecha", "Producto", "Cantidad"])

        registrar_evento(
            wb, modulo="Ventas despacho", tipo_evento="VENTA_DESPACHO", entidad="Tabla6",
            id_entidad=f"fila {fila}", accion="INSERT", estado=ORDEN_CORRECTA,
            valor_referencia=round(precio * cantidad, 2),
            detalle=f"{cantidad} x {nombre} a {precio} ({item['codigo']}, {item['origen']}{nota_descuento})",
        )
        return {"hubo_cambios": True, "estado": "REGISTRADA", "fila": fila, "producto": nombre,
                "codigo": item["codigo"], "cantidad": cantidad, "precio": precio,
                "total": round(precio * cantidad, 2), "fecha": fecha.date().isoformat(),
                "hora": hora.isoformat()}

    return editar_excel(_editar, descripcion=f"Venta despacho {cantidad} x {producto}")


def anular_venta_despacho(producto, cantidad, momento):
    """Cambia a 'ORDEN ANULADA' la venta que coincida con Fecha+Hora+Producto+Cantidad.

    Solo toca la columna 'Estado de orden'. Nunca cambia fecha, producto,
    cantidad ni precio de una venta ya escrita.
    """
    cantidad = _validar_cantidad(cantidad)
    fecha, hora = _separar_fecha_hora(momento)

    def _editar(wb, wb_valores):
        nombre = Catalogo(wb_valores).resolver(producto, permitir_combo_inactivo=True)["nombre"]
        ws, tabla = obtener_tabla(wb, HOJA, TABLA)
        venta = _buscar_venta(ws, tabla, fecha, hora, nombre, cantidad)
        if venta is None:
            raise LookupError(f"No existe una venta {cantidad} x {nombre} el {fecha.date()} a las {hora}.")
        if venta.get("Estado de orden") == ORDEN_ANULADA:
            return {"hubo_cambios": False, "estado": "YA_ESTABA_ANULADA", "fila": venta["_fila"]}

        letra_estado = columnas_de_tabla(tabla)["Estado de orden"]
        ws[f"{letra_estado}{venta['_fila']}"] = ORDEN_ANULADA
        registrar_evento(
            wb, modulo="Ventas despacho", tipo_evento="ANULACION", entidad="Tabla6",
            id_entidad=f"fila {venta['_fila']}", accion="UPDATE Estado de orden", estado=ORDEN_ANULADA,
            detalle=f"{cantidad} x {nombre} del {fecha.date()} {hora}",
        )
        return {"hubo_cambios": True, "estado": "ANULADA", "fila": venta["_fila"]}

    return editar_excel(_editar, descripcion=f"Anular venta despacho {cantidad} x {producto}")
