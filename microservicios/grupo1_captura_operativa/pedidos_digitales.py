"""Microservicio: Pedidos digitales = 'Interface digital' + 'Detalle comanda digital'.

FUNCION COMPARTIDA (Regla #4 y Regla #5 del libro de mapas): la usan el bot
de WhatsApp y, en el futuro, la web del SaaS. Las reglas de negocio (precios,
descuentos, envio) viven SOLO aqui, en un unico lugar.

Una comanda produce DOS escrituras en el mismo guardado:
  1. 'Interface digital'       -> 1 fila resumen del pedido
  2. 'Detalle comanda digital' -> 1 fila por cada linea del pedido
Todo se escribe como VALORES FIJOS (nunca formulas vivas): si mañana cambia
un precio, los pedidos ya cobrados no cambian.

Decisiones confirmadas por el usuario:
  - Combo = UNA sola linea (igual que en Ventas despacho). (23-sep-2026)
  - Solo se venden combos ACTIVOS, en despacho y en linea. (23-sep-2026)

Como se calculan los montos (sin doble descuento):
  - Linea de COMBO: precio unitario = 'Precio de Oferta' de Tabla Combos
    Dinamicos, que YA TRAE el descuento del combo adentro. DESCUENTO_ASIGNADO
    muestra el % del combo solo como trazabilidad; no se vuelve a restar.
  - Linea INDIVIDUAL: precio de TABLA CRUCE MAESTRA. Si NODO INTEGRAL marca
    Feriado o Dia de Pago activo hoy, se aplica el % de 'Interface digital'!B5
    (5%) a TODOS los productos. Esos dias no hay combos (los rechaza el catalogo).
  - SUBTOTAL_BRUTO  = suma de (cantidad x precio unitario)
  - DESCUENTO_TOTAL = suma de los descuentos de feriado / dia de pago
  - ENVIO_MAPS      = DISTANCIA_KM x Tarifa Base KM (NODO INTEGRAL)
  - TOTAL_A_COBRAR  = SUBTOTAL_BRUTO - DESCUENTO_TOTAL + ENVIO_MAPS

Duplicados: cada pedido trae un 'id_mensaje' unico (el ID del mensaje de
WhatsApp que confirmo el pedido). Queda anotado en RegistroEventos; si el
mismo id llega otra vez (reintento del webhook), no se escribe nada.
"""
from datetime import datetime, time

from conexion_drive import editar_excel
from conexion_drive.tablas import (
    agregar_fila, columnas_de_tabla, leer_tabla, obtener_tabla, registrar_evento,
)
from conexion_drive.tiempo import a_hora_local, ahora_local

from .catalogo import Catalogo

HOJA_INTERFACE, TABLA_INTERFACE = "Interface digital", "TablaInterfaceDigital"
HOJA_DETALLE, TABLA_DETALLE = "Detalle comanda digital", "Tabla3"
CONFIRMADO, ANULADO = "CONFIRMADO", "ANULADO"
TIPO_EVENTO = "PEDIDO_DIGITAL"


# ----------------------------------------------------------------------------
# Calculo del pedido (sin tocar el Excel todavia)
# ----------------------------------------------------------------------------
def _normalizar_lineas(lineas):
    """Acepta el formato del bot ({'producto_id', 'cantidad'}) o el del SaaS ({'producto', 'cantidad'})."""
    if not lineas:
        raise ValueError("El pedido no tiene productos.")
    normalizadas = []
    for linea in lineas:
        producto = linea.get("producto") or linea.get("producto_id")
        cantidad = linea.get("cantidad")
        if isinstance(cantidad, bool) or not isinstance(cantidad, int) or cantidad <= 0:
            raise ValueError(f"Cantidad invalida para '{producto}': {cantidad!r}")
        normalizadas.append({"producto": producto, "cantidad": cantidad})
    return normalizadas


def calcular_pedido(lineas, catalogo, distancia_km=None):
    """Devuelve las lineas calculadas y los 4 montos del pedido."""
    reglas = catalogo.reglas
    calculadas = []
    for linea in _normalizar_lineas(lineas):
        item = catalogo.resolver(linea["producto"])       # rechaza combos inactivos y nombres inexistentes
        cantidad = linea["cantidad"]
        precio = round(float(item["precio"]), 2)
        bruto = round(cantidad * precio, 2)

        if item["origen"] == "Tabla Combos Dinamicos":
            pct_combo = item.get("descuento_pct")
            pct_combo = float(pct_combo) / 100 if isinstance(pct_combo, (int, float)) and pct_combo > 1 else (pct_combo or 0)
            codigo_oferta, detonador, pct_asignado, descuento = item["codigo"], item["tipo_detonador"] or "Ninguno", pct_combo, 0.0
        elif reglas["descuento_general"]:
            codigo_oferta, pct_asignado = None, reglas["pct_general"]
            detonador = reglas["detonador_general"]
            descuento = round(bruto * pct_asignado, 2)
        else:
            codigo_oferta, detonador, pct_asignado, descuento = None, "Ninguno", 0.0, 0.0

        calculadas.append({
            "nombre": item["nombre"], "codigo": item["codigo"], "cantidad": cantidad,
            "precio": precio, "bruto": bruto, "descuento": descuento,
            "subtotal": round(bruto - descuento, 2), "codigo_oferta": codigo_oferta,
            "detonador": detonador, "pct": pct_asignado,
        })

    subtotal_bruto = round(sum(l["bruto"] for l in calculadas), 2)
    descuento_total = round(sum(l["descuento"] for l in calculadas), 2)
    envio = round(float(distancia_km) * reglas["tarifa_km"], 2) if distancia_km is not None else 0.0
    return {
        "lineas": calculadas,
        "subtotal_bruto": subtotal_bruto,
        "descuento_total": descuento_total,
        "envio": envio,
        "total": round(subtotal_bruto - descuento_total + envio, 2),
    }


# ----------------------------------------------------------------------------
# Ayudas de escritura
# ----------------------------------------------------------------------------
def _buscar_pedido_por_mensaje(wb, id_mensaje):
    ws, tabla = obtener_tabla(wb, "RegistroEventos", "tblRegistroEventos")
    marca = f"id_mensaje={id_mensaje} "
    for evento in leer_tabla(ws, tabla):
        if evento.get("TIPO_EVENTO") == TIPO_EVENTO and str(evento.get("DETALLE") or "").startswith(marca):
            return evento.get("ID_ENTIDAD")
    return None


def _siguiente_id_pedido(ws, tabla, prefijo):
    numeros = []
    for fila in leer_tabla(ws, tabla):
        texto = str(fila.get("ID_PEDIDO") or "")
        if texto.startswith(prefijo) and texto[len(prefijo):].isdigit():
            numeros.append(int(texto[len(prefijo):]))
    return f"{prefijo}{max(numeros, default=100) + 1}"


def _link_ubicacion(ubicacion):
    if not ubicacion:
        return None
    lat, lng = ubicacion
    return f"https://maps.google.com/?q={lat},{lng}"


# ----------------------------------------------------------------------------
# Funciones publicas
# ----------------------------------------------------------------------------
def registrar_pedido_digital(lineas, id_mensaje, cliente=None, metodo_pago=None,
                             ubicacion=None, distancia_km=None, momento=None,
                             fuente="WhatsApp", prefijo_id="WA-"):
    """Registra un pedido confirmado en 'Interface digital' + 'Detalle comanda digital'.

    lineas:       [{'producto': 'P01', 'cantidad': 3}, {'producto': 'co02', 'cantidad': 1}]
                  (codigo o nombre exacto; tambien acepta la clave 'producto_id' del bot)
    id_mensaje:   ID unico del mensaje que confirmo el pedido (evita duplicados)
    ubicacion:    (latitud, longitud) compartida por el cliente, o None
    distancia_km: distancia calculada por Google Maps (Grupo 3). Mientras no
                  exista esa integracion, llega None y el envio queda en 0.
    """
    if not id_mensaje:
        raise ValueError("Falta id_mensaje: sin el no se pueden detectar pedidos duplicados.")
    momento = a_hora_local(momento) if momento else ahora_local()
    fecha = datetime(momento.year, momento.month, momento.day)
    hora = time(momento.hour, momento.minute, momento.second)

    def _editar(wb, wb_valores):
        repetido = _buscar_pedido_por_mensaje(wb, id_mensaje)
        if repetido:
            return {"hubo_cambios": False, "estado": "DUPLICADO", "id_pedido": repetido}

        pedido = calcular_pedido(lineas, Catalogo(wb_valores), distancia_km)
        ws_i, tabla_i = obtener_tabla(wb, HOJA_INTERFACE, TABLA_INTERFACE)
        ws_d, tabla_d = obtener_tabla(wb, HOJA_DETALLE, TABLA_DETALLE)
        id_pedido = _siguiente_id_pedido(ws_i, tabla_i, prefijo_id)
        codigos_oferta = sorted({l["codigo_oferta"] for l in pedido["lineas"] if l["codigo_oferta"]})

        fila_resumen = agregar_fila(ws_i, tabla_i, {
            "ID_PEDIDO": id_pedido,
            "FECHA": fecha,
            "HORA": hora,
            "FUENTE": fuente,
            "CLIENTE": cliente,
            "DETALLE_COMANDA": ", ".join(f"{l['cantidad']} {l['nombre']}" for l in pedido["lineas"]),
            "SUBTOTAL_BRUTO": pedido["subtotal_bruto"],
            "DESCUENTO_TOTAL": pedido["descuento_total"],
            "ENVIO_MAPS": pedido["envio"],
            "TOTAL_A_COBRAR": pedido["total"],
            "METODO_PAGO": metodo_pago,
            "LINK_UBICACION": _link_ubicacion(ubicacion),
            "ESTADO": CONFIRMADO,
            "DISTANCIA_KM": distancia_km,
            "CODIGO_OFERTA": ", ".join(codigos_oferta) or None,
        }, columnas_clave=["ID_PEDIDO"])

        filas_detalle = []
        for linea in pedido["lineas"]:
            filas_detalle.append(agregar_fila(ws_d, tabla_d, {
                "ID_PEDIDO": id_pedido,
                "FECHA": fecha,
                "PRODUCTO": linea["nombre"],
                "CANTIDAD": linea["cantidad"],
                "ESTADO": CONFIRMADO,
                "SUBTOTAL_LINEA": linea["subtotal"],
                "CODIGO_OFERTA": linea["codigo_oferta"],
                "TIPO_DETONADOR": linea["detonador"],
                "DESCUENTO_ASIGNADO": linea["pct"],
            }, columnas_clave=["ID_PEDIDO"]))

        registrar_evento(
            wb, modulo="Pedidos digitales", tipo_evento=TIPO_EVENTO, entidad="TablaInterfaceDigital",
            id_entidad=id_pedido, accion="INSERT", estado=CONFIRMADO, origen=fuente,
            valor_referencia=pedido["total"],
            detalle=f"id_mensaje={id_mensaje} | {len(pedido['lineas'])} lineas | "
                    + ("sin distancia (envio 0)" if distancia_km is None else f"{distancia_km} km"),
        )
        return {"hubo_cambios": True, "estado": "REGISTRADO", "id_pedido": id_pedido,
                "fila_resumen": fila_resumen, "filas_detalle": filas_detalle,
                "subtotal_bruto": pedido["subtotal_bruto"], "descuento_total": pedido["descuento_total"],
                "envio": pedido["envio"], "total": pedido["total"],
                "lineas": [(l["cantidad"], l["nombre"], l["subtotal"]) for l in pedido["lineas"]]}

    return editar_excel(_editar, descripcion=f"Pedido digital {id_mensaje}")


def anular_linea_pedido_digital(id_pedido, producto):
    """Anula UNA linea de un pedido (ESTADO -> 'ANULADO' en Detalle comanda digital).

    Segun el mapa, solo cambia esa linea: las demas lineas y la fila de
    'Interface digital' no se tocan.
    """
    def _editar(wb, wb_valores):
        nombre = Catalogo(wb_valores).resolver(producto, permitir_combo_inactivo=True)["nombre"]
        ws, tabla = obtener_tabla(wb, HOJA_DETALLE, TABLA_DETALLE)
        letra_estado = columnas_de_tabla(tabla)["ESTADO"]
        for linea in leer_tabla(ws, tabla):
            if linea.get("ID_PEDIDO") == id_pedido and linea.get("PRODUCTO") == nombre:
                if linea.get("ESTADO") == ANULADO:
                    continue
                ws[f"{letra_estado}{linea['_fila']}"] = ANULADO
                registrar_evento(
                    wb, modulo="Pedidos digitales", tipo_evento="ANULACION_LINEA",
                    entidad="Tabla3", id_entidad=id_pedido, accion="UPDATE ESTADO", estado=ANULADO,
                    detalle=f"{linea.get('CANTIDAD')} x {nombre} (fila {linea['_fila']})",
                )
                return {"hubo_cambios": True, "estado": "ANULADA", "fila": linea["_fila"]}
        raise LookupError(f"No hay una linea CONFIRMADA de '{nombre}' en el pedido {id_pedido}.")

    return editar_excel(_editar, descripcion=f"Anular {producto} en {id_pedido}")
