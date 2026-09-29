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
  - ENVIO_MAPS      = DISTANCIA_KM x Tarifa Base KM (NODO INTEGRAL C3), sin IVA,
                      redondeado HACIA ARRIBA al multiplo de 50 (regla unica,
                      vive en google_maps.calcular_costo_envio)
  - TOTAL_A_COBRAR  = SUBTOTAL_BRUTO - DESCUENTO_TOTAL + ENVIO_MAPS

Envio (Grupo 3B, 24-sep-2026). Este archivo es la UNICA puerta a Google Maps
para bot y SaaS. El cliente debe ver el TOTAL COMPLETO antes de confirmar:

  1. El cliente comparte su ubicacion:
         cot = cotizar_envio((lat, lng))
     Si cot['dentro_zona'] es False, el bot avisa (fuera del radio de entrega).
     El bot guarda cot['km'] y cot['minutos'] en el estado de la conversacion.
  2. Antes de confirmar, el bot muestra el total completo:
         resumen = cotizar_pedido(lineas, distancia_km=cot['km'], minutos_entrega=cot['minutos'])
  3. El cliente dice "Si":
         registrar_pedido_digital(lineas, id_mensaje, ubicacion=(lat, lng),
                                  distancia_km=cot['km'], minutos_entrega=cot['minutos'],
                                  total_cotizado=resumen['total'])
     Como ya trae los km, NO se vuelve a llamar a Google (no se paga dos veces).
     Si el total cambio entre la cotizacion y el "Si" (ej. se apago un combo),
     se rechaza sin escribir nada y el bot vuelve a cotizar.

Sin ubicacion ni km = retiro en el local: envio 0 y NODO C9/C10 no se tocan.
Fuera de zona o Google caido = se rechaza y no se escribe nada.

Duplicados: cada pedido trae un 'id_mensaje' unico (el ID del mensaje de
WhatsApp que confirmo el pedido). Queda anotado en RegistroEventos; si el
mismo id llega otra vez (reintento del webhook), no se escribe nada.

Forma de pago (decision del usuario, 27-sep-2026): el cliente elige en el bot
"Efectivo" o "SINPE Movil" y la forma de pago llega JUNTO con la comanda.
Es OBLIGATORIA: un pedido sin forma de pago (o con otra) se rechaza sin
escribir nada. Las lineas solo se pueden quitar en la conversacion, ANTES
de "Pedido confirmado"; despues ya no se tocan lineas sueltas.

Anulacion (#7, Regla #30, 27-sep-2026) -- solo la ADMINISTRACION (SaaS), el
bot nunca anula. Una comanda es una unidad: "si pide 3 articulos paga 3".
    anular_pedido_digital('WA-119', 'El cliente no lo recibio')
  - 'Interface digital': fila NUEVA 'WA-119-A' con los montos en NEGATIVO
    (comanda + envio), ESTADO "CONFIRMADO" y la FECHA del pedido original.
    Asi 'Flujo caja diario' (SUMIFS por fecha con CONFIRMADO) resta el pedido
    ese mismo dia en "Ventas digitales" y en "Envios cobrados".
  - 'Detalle comanda digital': filas de RASTRO 'WA-119-A' con ESTADO
    "ANULADO". NO devuelven inventario: 'Control de Inventario' E solo suma
    las lineas CONFIRMADO, y la linea original sigue CONFIRMADO (los
    alimentos que salen de la panaderia no se reciben de vuelta).
  - El registro original NO se modifica (queda el rastro contable).
  - Se rechaza (sin escribir nada): pedido pagado por SINPE Movil (pago en
    firme), pedido inexistente, pedido ya anulado, una fila de anulacion o
    un motivo vacio.
  - V m013 (29-sep-2026): tampoco se anula un pedido YA ENTREGADO (el cliente
    recibio el pan y pago).

Telefono y entrega (Sesion A del bot, 29-sep-2026; requiere la migracion m013):
  - #33 opcion B: registrar_pedido_digital(..., telefono='whatsapp:+506...') guarda
    el numero LIMPIO (+506...) en la columna TELEFONO. Es opcional (el SaaS puede
    no traerlo); el bot siempre lo manda. El cliente lo acepta en el mensaje #28.
  - Recibo al entregar, opcion C: marcar_entregado('WA-120') anota la fecha y hora
    en la columna ENTREGADO y devuelve los DATOS DEL RECIBO. No envia WhatsApp: el
    bot envia el recibo (hoy por el comando 'ENTREGADO WA-120' de un numero
    autorizado; despues, el boton de administracion del SaaS llama esta misma
    funcion). ESTADO no cambia (sigue CONFIRMADO): 'Flujo caja diario' y 'Control
    de Inventario' suman solo CONFIRMADO. Si ya estaba entregado, no escribe nada
    (estado YA_ENTREGADO) y el cliente no recibe dos recibos.
"""
import unicodedata
from datetime import datetime, time

from conexion_drive import editar_excel
from conexion_drive.tablas import (
    agregar_fila, columnas_de_tabla, leer_tabla, obtener_tabla, registrar_evento,
)
from conexion_drive.tiempo import a_hora_local, ahora_local

from conexion_drive import leer_excel
from microservicios.grupo3_integraciones import google_maps
from microservicios.grupo3_integraciones.google_maps import ErrorGoogleMaps  # noqa: F401 (lo usan bot y pruebas)

from .catalogo import Catalogo, leer_reglas_del_dia

HOJA_INTERFACE, TABLA_INTERFACE = "Interface digital", "TablaInterfaceDigital"
HOJA_DETALLE, TABLA_DETALLE = "Detalle comanda digital", "Tabla3"
CONFIRMADO, ANULADO = "CONFIRMADO", "ANULADO"
TIPO_EVENTO = "PEDIDO_DIGITAL"
HOJA_NODO = "NODO INTEGRAL"
ETIQUETA_MINUTOS, ETIQUETA_KM = "Tiempo Estimado Delivery", "Distancia Delivery (KM)"
EFECTIVO, SINPE = "Efectivo", "SINPE M\u00f3vil"          # los 2 valores que se escriben en METODO_PAGO
SUFIJO_ANULACION = "-A"                                   # WA-119 -> WA-119-A
TIPO_EVENTO_ANULACION = "ANULACION_PEDIDO"
TIPO_EVENTO_ENTREGA = "ENTREGA_PEDIDO"
COL_TELEFONO, COL_ENTREGADO = "TELEFONO", "ENTREGADO"     # columnas nuevas de la m013
FORMATO_ENTREGADO = "d-mmm hh:mm"


class FueraDeZona(Exception):
    """El cliente esta mas lejos que el radio de entrega (RADIO_MAXIMO_KM)."""

    def __init__(self, km, radio_maximo_km):
        self.km, self.radio_maximo_km = km, radio_maximo_km
        super().__init__(f"Fuera de zona: {km} km es mas que el radio de entrega de "
                         f"{radio_maximo_km:g} km. No se escribio nada.")


class TotalCambio(Exception):
    """El total a registrar no coincide con el que se le mostro al cliente."""


class MetodoPagoInvalido(ValueError):
    """Falta la forma de pago o no es 'Efectivo' / 'SINPE Movil'."""


class AnulacionNoPermitida(Exception):
    """La anulacion va contra una regla del negocio (no se escribe nada)."""


class EntregaNoPermitida(Exception):
    """El pedido no se puede marcar como entregado (no se escribe nada)."""


class TelefonoInvalido(ValueError):
    """El telefono no tiene forma de numero (+ y digitos)."""


# ----------------------------------------------------------------------------
# Telefono (#33 opcion B)
# ----------------------------------------------------------------------------
def normalizar_telefono(valor):
    """'whatsapp:+506 8888-7777' -> '+50688887777'. Vacio -> None. Otra cosa -> TelefonoInvalido."""
    texto = str(valor or "").strip()
    if texto.lower().startswith("whatsapp:"):
        texto = texto[len("whatsapp:"):]
    texto = "".join(texto.split()).replace("-", "")
    if not texto:
        return None
    cuerpo = texto[1:] if texto.startswith("+") else texto
    if not cuerpo.isdigit() or not 8 <= len(cuerpo) <= 15:
        raise TelefonoInvalido(f"Telefono invalido: {valor!r}. No se escribio nada.")
    return "+" + cuerpo if texto.startswith("+") else cuerpo


# ----------------------------------------------------------------------------
# Forma de pago
# ----------------------------------------------------------------------------
def _texto_simple(valor):
    """'  SINPE  Móvil ' -> 'sinpe movil' (sin tildes, minusculas, un espacio)."""
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


def normalizar_metodo_pago(valor):
    """Devuelve 'Efectivo' o 'SINPE Móvil' (escritura unica). Si no, MetodoPagoInvalido."""
    texto = _texto_simple(valor)
    if texto == "efectivo":
        return EFECTIVO
    if texto in ("sinpe", "sinpe movil"):
        return SINPE
    raise MetodoPagoInvalido(
        f"Forma de pago invalida: {valor!r}. Debe ser '{EFECTIVO}' o '{SINPE}'. No se escribio nada."
    )


def _es_sinpe(valor):
    """Para filas YA escritas: vacio o 'Efectivo' = no es SINPE (pedidos viejos de prueba)."""
    return _texto_simple(valor).startswith("sinpe")


# ----------------------------------------------------------------------------
# Envio (Google Maps)
# ----------------------------------------------------------------------------
def _resolver_envio(ubicacion, distancia_km, minutos_entrega, tarifa_km):
    """Decide km, minutos y costo de envio. Lanza FueraDeZona o ErrorGoogleMaps.

    Tres caminos:
      - km ya cotizados        -> se usan tal cual (NO se llama a Google)
      - solo ubicacion         -> se pregunta a Google Maps
      - ni ubicacion ni km     -> retiro en el local, envio 0
    """
    if distancia_km is None and ubicacion is None:
        return {"km": None, "minutos": None, "envio": 0, "origen": "RETIRO EN LOCAL"}

    if tarifa_km <= 0:
        raise ErrorGoogleMaps("NODO INTEGRAL 'Tarifa Base KM' esta vacia o en 0: no se puede cobrar envio.")
    cfg = google_maps.configuracion()

    if distancia_km is None:
        lat, lng = ubicacion
        ruta = google_maps.consultar_ruta(lat, lng)
        km, minutos, origen = ruta["km"], ruta["minutos"], "GOOGLE MAPS"
    else:
        if isinstance(distancia_km, bool) or not isinstance(distancia_km, (int, float)) or distancia_km < 0:
            raise ValueError(f"distancia_km invalida: {distancia_km!r}")
        km, minutos, origen = round(float(distancia_km), 2), minutos_entrega, "COTIZACION"

    if km > cfg["radio_maximo_km"]:
        raise FueraDeZona(km, cfg["radio_maximo_km"])
    envio = google_maps.calcular_costo_envio(km, tarifa_km, cfg["redondeo"])
    return {"km": km, "minutos": minutos, "envio": envio, "origen": origen}


def _escribir_ultimo_envio_en_nodo(wb, km, minutos):
    """NODO INTEGRAL C9 (minutos) y C10 (km) = solo el ULTIMO pedido (referencia)."""
    ws = wb[HOJA_NODO]
    for fila in range(1, ws.max_row + 1):
        etiqueta = ws.cell(fila, 2).value
        if not isinstance(etiqueta, str):
            continue
        if etiqueta.strip() == ETIQUETA_MINUTOS and minutos is not None:
            ws.cell(fila, 3).value = minutos
        elif etiqueta.strip() == ETIQUETA_KM:
            ws.cell(fila, 3).value = km


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


def calcular_pedido(lineas, catalogo, envio=0):
    """Devuelve las lineas calculadas y los 4 montos del pedido.

    envio: colones ya calculados por _resolver_envio (la regla de redondeo
    vive en un solo lugar: google_maps.calcular_costo_envio).
    """
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
def cotizar_envio(ubicacion):
    """SOLO LEE. El bot la llama cuando el cliente comparte su ubicacion.

    Devuelve: dentro_zona, km, minutos, envio (None si fuera de zona),
    radio_maximo_km, link_ubicacion. Si Google falla lanza ErrorGoogleMaps.
    """
    if not ubicacion:
        raise ValueError("cotizar_envio necesita la ubicacion (latitud, longitud).")
    tarifa = leer_reglas_del_dia(leer_excel())["tarifa_km"]
    lat, lng = ubicacion
    try:
        info = _resolver_envio(ubicacion, None, None, tarifa)
        return {"dentro_zona": True, "km": info["km"], "minutos": info["minutos"], "envio": info["envio"],
                "radio_maximo_km": google_maps.configuracion()["radio_maximo_km"],
                "link_ubicacion": _link_ubicacion((lat, lng))}
    except FueraDeZona as fuera:
        return {"dentro_zona": False, "km": fuera.km, "minutos": None, "envio": None,
                "radio_maximo_km": fuera.radio_maximo_km, "link_ubicacion": _link_ubicacion((lat, lng))}


def cotizar_pedido(lineas, ubicacion=None, distancia_km=None, minutos_entrega=None):
    """SOLO LEE. Total COMPLETO (productos - descuentos + envio) para mostrarlo
    al cliente ANTES de que confirme. Usa exactamente el mismo calculo que
    registrar_pedido_digital, asi lo que ve el cliente es lo que se cobra.

    Rechaza (lanza error) productos inexistentes, combos inactivos, fuera de
    zona (FueraDeZona) y fallas de Google (ErrorGoogleMaps).
    """
    catalogo = Catalogo(leer_excel())
    info = _resolver_envio(ubicacion, distancia_km, minutos_entrega, catalogo.reglas["tarifa_km"])
    pedido = calcular_pedido(lineas, catalogo, info["envio"])
    return {"subtotal_bruto": pedido["subtotal_bruto"], "descuento_total": pedido["descuento_total"],
            "envio": pedido["envio"], "total": pedido["total"],
            "km": info["km"], "minutos": info["minutos"], "origen_envio": info["origen"],
            "lineas": [(l["cantidad"], l["nombre"], l["subtotal"]) for l in pedido["lineas"]]}


def registrar_pedido_digital(lineas, id_mensaje, cliente=None, metodo_pago=None,
                             ubicacion=None, distancia_km=None, minutos_entrega=None,
                             total_cotizado=None, momento=None,
                             fuente="WhatsApp", prefijo_id="WA-", telefono=None):
    """Registra un pedido confirmado en 'Interface digital' + 'Detalle comanda digital'.

    lineas:       [{'producto': 'P01', 'cantidad': 3}, {'producto': 'co02', 'cantidad': 1}]
                  (codigo o nombre exacto; tambien acepta la clave 'producto_id' del bot)
    id_mensaje:   ID unico del mensaje que confirmo el pedido (evita duplicados)
    ubicacion:    (latitud, longitud) compartida por el cliente, o None
    distancia_km, minutos_entrega: los que devolvio cotizar_envio(). Si llegan,
                  NO se vuelve a llamar a Google. Si solo llega 'ubicacion', se
                  consulta Google aqui (despues de revisar duplicados).
                  Sin ubicacion ni km = retiro en el local (envio 0).
    total_cotizado: el total que se le mostro al cliente (cotizar_pedido). Si
                  ya no coincide, se lanza TotalCambio y no se escribe nada.
    metodo_pago:  OBLIGATORIO, 'Efectivo' o 'SINPE Movil' (lo elige el cliente en
                  el bot y llega con la comanda). Si falta: MetodoPagoInvalido.
    telefono:     OPCIONAL (m013, #33 opcion B). El del cliente, tal como lo manda
                  Twilio ('whatsapp:+506...'); se guarda limpio en TELEFONO.
    """
    if not id_mensaje:
        raise ValueError("Falta id_mensaje: sin el no se pueden detectar pedidos duplicados.")
    metodo_pago = normalizar_metodo_pago(metodo_pago)    # rechaza antes de abrir el Excel
    telefono = normalizar_telefono(telefono)             # idem
    momento = a_hora_local(momento) if momento else ahora_local()
    fecha = datetime(momento.year, momento.month, momento.day)
    hora = time(momento.hour, momento.minute, momento.second)

    def _editar(wb, wb_valores):
        repetido = _buscar_pedido_por_mensaje(wb, id_mensaje)
        if repetido:
            return {"hubo_cambios": False, "estado": "DUPLICADO", "id_pedido": repetido}

        catalogo = Catalogo(wb_valores)
        info = _resolver_envio(ubicacion, distancia_km, minutos_entrega, catalogo.reglas["tarifa_km"])
        pedido = calcular_pedido(lineas, catalogo, info["envio"])
        if total_cotizado is not None and round(float(total_cotizado), 2) != pedido["total"]:
            raise TotalCambio(
                f"El total cambio desde la cotizacion: se mostro {total_cotizado} y ahora es "
                f"{pedido['total']}. No se escribio nada; hay que volver a cotizar."
            )
        ws_i, tabla_i = obtener_tabla(wb, HOJA_INTERFACE, TABLA_INTERFACE)
        ws_d, tabla_d = obtener_tabla(wb, HOJA_DETALLE, TABLA_DETALLE)
        id_pedido = _siguiente_id_pedido(ws_i, tabla_i, prefijo_id)
        codigos_oferta = sorted({l["codigo_oferta"] for l in pedido["lineas"] if l["codigo_oferta"]})

        valores_resumen = {
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
            "DISTANCIA_KM": info["km"],
            "CODIGO_OFERTA": ", ".join(codigos_oferta) or None,
        }
        if telefono:                     # solo si llega (asi sirve tambien con copias sin la m013)
            valores_resumen[COL_TELEFONO] = telefono
        fila_resumen = agregar_fila(ws_i, tabla_i, valores_resumen, columnas_clave=["ID_PEDIDO"])

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

        if info["km"] is not None:
            _escribir_ultimo_envio_en_nodo(wb, info["km"], info["minutos"])

        if info["km"] is None:
            texto_envio = "retiro en local (envio 0)"
        else:
            texto_envio = (f"{info['km']} km" + (f", {info['minutos']} min" if info["minutos"] is not None else "")
                           + f", envio {info['envio']} ({info['origen']})")
        registrar_evento(
            wb, modulo="Pedidos digitales", tipo_evento=TIPO_EVENTO, entidad="TablaInterfaceDigital",
            id_entidad=id_pedido, accion="INSERT", estado=CONFIRMADO, origen=fuente,
            valor_referencia=pedido["total"],
            detalle=f"id_mensaje={id_mensaje} | {len(pedido['lineas'])} lineas | {texto_envio}",
        )
        return {"hubo_cambios": True, "estado": "REGISTRADO", "id_pedido": id_pedido,
                "fila_resumen": fila_resumen, "filas_detalle": filas_detalle,
                "subtotal_bruto": pedido["subtotal_bruto"], "descuento_total": pedido["descuento_total"],
                "envio": pedido["envio"], "total": pedido["total"],
                "metodo_pago": metodo_pago,          # para que la comanda diga como paga el cliente
                "telefono": telefono,
                "km": info["km"], "minutos": info["minutos"], "origen_envio": info["origen"],
                "lineas": [(l["cantidad"], l["nombre"], l["subtotal"]) for l in pedido["lineas"]]}

    return editar_excel(_editar, descripcion=f"Pedido digital {id_mensaje}")


def _numero(valor, columna, id_pedido):
    """Monto de una celda: vacio = 0; texto u otra cosa = error (no se adivina)."""
    if valor is None or valor == "":
        return 0.0
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise ValueError(f"{id_pedido}: la columna {columna} no tiene un numero ({valor!r}). No se escribio nada.")
    return float(valor)


def _negativo(monto):
    """-monto redondeado; el '+ 0.0' evita que un 0 quede escrito como -0.0."""
    return round(-monto, 2) + 0.0


def anular_pedido_digital(id_pedido, motivo, momento=None, origen="SaaS"):
    """Anula la COMANDA COMPLETA (solo la administracion). Ver el encabezado del archivo.

    id_pedido: el ID original, p. ej. 'WA-119'
    motivo:    texto obligatorio (queda en DETALLE_COMANDA y en RegistroEventos)
    """
    id_pedido = str(id_pedido or "").strip()
    motivo = " ".join(str(motivo or "").split())
    if not id_pedido:
        raise ValueError("Falta el id_pedido a anular.")
    if not motivo:
        raise AnulacionNoPermitida("Falta el motivo de la anulacion. No se escribio nada.")
    if id_pedido.endswith(SUFIJO_ANULACION):
        raise AnulacionNoPermitida(f"{id_pedido} es una fila de anulacion: no se puede anular. No se escribio nada.")
    id_anulacion = id_pedido + SUFIJO_ANULACION
    momento = a_hora_local(momento) if momento else ahora_local()
    hora = time(momento.hour, momento.minute, momento.second)

    def _editar(wb, wb_valores):
        ws_i, tabla_i = obtener_tabla(wb, HOJA_INTERFACE, TABLA_INTERFACE)
        ws_d, tabla_d = obtener_tabla(wb, HOJA_DETALLE, TABLA_DETALLE)
        col_i = columnas_de_tabla(tabla_i)
        col_d = columnas_de_tabla(tabla_d)
        val_i, val_d = wb_valores[HOJA_INTERFACE], wb_valores[HOJA_DETALLE]   # montos ya calculados

        filas_i = leer_tabla(ws_i, tabla_i)
        if any(str(f.get("ID_PEDIDO") or "").strip() == id_anulacion for f in filas_i):
            raise AnulacionNoPermitida(f"El pedido {id_pedido} ya estaba anulado ({id_anulacion}). No se escribio nada.")
        original = next((f for f in filas_i if str(f.get("ID_PEDIDO") or "").strip() == id_pedido), None)
        if original is None:
            raise LookupError(f"No existe el pedido {id_pedido} en '{HOJA_INTERFACE}'. No se escribio nada.")

        def valor_i(columna):
            return val_i[f"{col_i[columna]}{original['_fila']}"].value

        if valor_i("ESTADO") != CONFIRMADO:
            raise AnulacionNoPermitida(
                f"El pedido {id_pedido} no esta CONFIRMADO (estado {valor_i('ESTADO')!r}): no suma en el flujo, "
                f"no hay nada que anular. No se escribio nada.")
        metodo = valor_i("METODO_PAGO")
        if _es_sinpe(metodo):
            raise AnulacionNoPermitida(
                f"El pedido {id_pedido} se pago por {metodo}: el pago es en firme y no se anula (Regla #30). "
                f"No se escribio nada.")
        if COL_ENTREGADO in col_i and valor_i(COL_ENTREGADO) not in (None, ""):
            raise AnulacionNoPermitida(
                f"El pedido {id_pedido} ya fue ENTREGADO ({valor_i(COL_ENTREGADO)}): el cliente recibio el "
                f"pedido y pago; no se anula (decision del usuario, 29-sep-2026). No se escribio nada.")

        subtotal = _numero(valor_i("SUBTOTAL_BRUTO"), "SUBTOTAL_BRUTO", id_pedido)
        descuento = _numero(valor_i("DESCUENTO_TOTAL"), "DESCUENTO_TOTAL", id_pedido)
        envio = _numero(valor_i("ENVIO_MAPS"), "ENVIO_MAPS", id_pedido)
        total = _numero(valor_i("TOTAL_A_COBRAR"), "TOTAL_A_COBRAR", id_pedido)

        # Lineas CONFIRMADAS del pedido original (las que rebajaron inventario)
        lineas = []
        for linea in leer_tabla(ws_d, tabla_d):
            if str(linea.get("ID_PEDIDO") or "").strip() != id_pedido:
                continue
            fila = linea["_fila"]
            if val_d[f"{col_d['ESTADO']}{fila}"].value != CONFIRMADO:
                continue
            lineas.append({columna: val_d[f"{letra}{fila}"].value for columna, letra in col_d.items()})

        # 1) Interface digital: fila negativa, CONFIRMADO, fecha ORIGINAL
        fila_resumen = agregar_fila(ws_i, tabla_i, {
            "ID_PEDIDO": id_anulacion,
            "FECHA": valor_i("FECHA"),
            "HORA": hora,
            "FUENTE": valor_i("FUENTE"),
            "CLIENTE": valor_i("CLIENTE"),
            "DETALLE_COMANDA": f"ANULADO: {motivo} (anula {id_pedido})",
            "SUBTOTAL_BRUTO": _negativo(subtotal),
            "DESCUENTO_TOTAL": _negativo(descuento),
            "ENVIO_MAPS": _negativo(envio),
            "TOTAL_A_COBRAR": _negativo(total),
            "METODO_PAGO": metodo,
            "ESTADO": CONFIRMADO,
            "CODIGO_OFERTA": valor_i("CODIGO_OFERTA"),
        }, columnas_clave=["ID_PEDIDO"])

        # 2) Detalle comanda digital: rastro ANULADO (no devuelve inventario)
        filas_detalle = []
        for linea in lineas:
            filas_detalle.append(agregar_fila(ws_d, tabla_d, {
                "ID_PEDIDO": id_anulacion,
                "FECHA": linea.get("FECHA"),
                "PRODUCTO": linea.get("PRODUCTO"),
                "CANTIDAD": linea.get("CANTIDAD"),
                "ESTADO": ANULADO,
                "SUBTOTAL_LINEA": _negativo(_numero(linea.get("SUBTOTAL_LINEA"), "SUBTOTAL_LINEA", id_pedido)),
                "CODIGO_OFERTA": linea.get("CODIGO_OFERTA"),
                "TIPO_DETONADOR": linea.get("TIPO_DETONADOR"),
                "DESCUENTO_ASIGNADO": linea.get("DESCUENTO_ASIGNADO"),
            }, columnas_clave=["ID_PEDIDO"]))

        # 3) Evento
        registrar_evento(
            wb, modulo="Pedidos digitales", tipo_evento=TIPO_EVENTO_ANULACION, entidad="TablaInterfaceDigital",
            id_entidad=id_anulacion, accion="INSERT", estado=ANULADO, origen=origen,
            valor_referencia=_negativo(total),
            detalle=(f"anula={id_pedido} | motivo={motivo} | pago={metodo or 'sin dato'} | "
                     f"{len(filas_detalle)} lineas de rastro (sin devolver inventario)"),
        )
        return {"hubo_cambios": True, "estado": "ANULADO", "id_pedido": id_pedido,
                "id_anulacion": id_anulacion, "fila_resumen": fila_resumen, "filas_detalle": filas_detalle,
                "total_revertido": _negativo(total), "envio_revertido": _negativo(envio)}

    return editar_excel(_editar, descripcion=f"Anular pedido {id_pedido}")


def marcar_entregado(id_pedido, momento=None, origen="WhatsApp"):
    """Marca un pedido como ENTREGADO y devuelve los datos del recibo (m013, opcion C).

    Funcion COMPARTIDA: hoy la llama el bot (comando 'ENTREGADO WA-120' de un numero
    autorizado) y despues la llamara el boton de administracion del SaaS. NO envia
    WhatsApp: quien la llama envia el recibo con los datos que devuelve.

    - Escribe SOLO la fecha y hora en la columna ENTREGADO (ESTADO no cambia) y un
      evento ENTREGA_PEDIDO en RegistroEventos.
    - Si ya estaba entregado: no escribe nada, estado 'YA_ENTREGADO'.
    - Rechaza (sin escribir nada): pedido inexistente (LookupError), fila de
      anulacion, pedido anulado o no CONFIRMADO (EntregaNoPermitida).
    """
    id_pedido = str(id_pedido or "").strip().upper()
    if not id_pedido:
        raise ValueError("Falta el id_pedido a marcar como entregado.")
    if id_pedido.endswith(SUFIJO_ANULACION):
        raise EntregaNoPermitida(f"{id_pedido} es una fila de anulacion: no se entrega. No se escribio nada.")
    momento = a_hora_local(momento) if momento else ahora_local()
    entregado_en = datetime(momento.year, momento.month, momento.day, momento.hour, momento.minute, momento.second)

    def _editar(wb, wb_valores):
        ws_i, tabla_i = obtener_tabla(wb, HOJA_INTERFACE, TABLA_INTERFACE)
        ws_d, tabla_d = obtener_tabla(wb, HOJA_DETALLE, TABLA_DETALLE)
        col_i, col_d = columnas_de_tabla(tabla_i), columnas_de_tabla(tabla_d)
        if COL_ENTREGADO not in col_i or COL_TELEFONO not in col_i:
            raise EntregaNoPermitida(
                f"'{HOJA_INTERFACE}' no tiene las columnas {COL_TELEFONO}/{COL_ENTREGADO}: falta aplicar la "
                "migracion m013. No se escribio nada.")
        val_i, val_d = wb_valores[HOJA_INTERFACE], wb_valores[HOJA_DETALLE]

        filas_i = leer_tabla(ws_i, tabla_i)
        original = next((f for f in filas_i if str(f.get("ID_PEDIDO") or "").strip().upper() == id_pedido), None)
        if original is None:
            raise LookupError(f"No existe el pedido {id_pedido} en '{HOJA_INTERFACE}'. No se escribio nada.")
        if any(str(f.get("ID_PEDIDO") or "").strip().upper() == id_pedido + SUFIJO_ANULACION for f in filas_i):
            raise EntregaNoPermitida(f"El pedido {id_pedido} fue ANULADO: no se entrega. No se escribio nada.")
        fila = original["_fila"]

        def valor_i(columna):
            return val_i[f"{col_i[columna]}{fila}"].value

        if valor_i("ESTADO") != CONFIRMADO:
            raise EntregaNoPermitida(
                f"El pedido {id_pedido} no esta CONFIRMADO (estado {valor_i('ESTADO')!r}). No se escribio nada.")

        lineas = []
        for linea in leer_tabla(ws_d, tabla_d):
            f = linea["_fila"]
            if (str(linea.get("ID_PEDIDO") or "").strip().upper() == id_pedido
                    and val_d[f"{col_d['ESTADO']}{f}"].value == CONFIRMADO):
                lineas.append((val_d[f"{col_d['CANTIDAD']}{f}"].value, val_d[f"{col_d['PRODUCTO']}{f}"].value,
                               _numero(val_d[f"{col_d['SUBTOTAL_LINEA']}{f}"].value, "SUBTOTAL_LINEA", id_pedido)))
        recibo = {
            "id_pedido": id_pedido,
            "cliente": valor_i("CLIENTE"),
            "telefono": valor_i(COL_TELEFONO),
            "fecha": valor_i("FECHA"),
            "lineas": lineas,
            "subtotal_bruto": _numero(valor_i("SUBTOTAL_BRUTO"), "SUBTOTAL_BRUTO", id_pedido),
            "descuento_total": _numero(valor_i("DESCUENTO_TOTAL"), "DESCUENTO_TOTAL", id_pedido),
            "envio": _numero(valor_i("ENVIO_MAPS"), "ENVIO_MAPS", id_pedido),
            "total": _numero(valor_i("TOTAL_A_COBRAR"), "TOTAL_A_COBRAR", id_pedido),
            "metodo_pago": valor_i("METODO_PAGO"),
        }

        celda = ws_i[f"{col_i[COL_ENTREGADO]}{fila}"]
        if celda.value not in (None, ""):
            return {"hubo_cambios": False, "estado": "YA_ENTREGADO", "entregado_en": celda.value, **recibo}
        celda.value = entregado_en
        celda.number_format = FORMATO_ENTREGADO
        registrar_evento(
            wb, modulo="Pedidos digitales", tipo_evento=TIPO_EVENTO_ENTREGA, entidad="TablaInterfaceDigital",
            id_entidad=id_pedido, accion="UPDATE", estado="ENTREGADO", origen=origen,
            valor_referencia=recibo["total"],
            detalle=(f"entregado={entregado_en:%Y-%m-%d %H:%M:%S} | pago={recibo['metodo_pago'] or 'sin dato'} | "
                     f"recibo={'si (hay telefono)' if recibo['telefono'] else 'no (sin telefono)'}"),
        )
        return {"hubo_cambios": True, "estado": "ENTREGADO", "entregado_en": entregado_en, **recibo}

    return editar_excel(_editar, descripcion=f"Entregar pedido {id_pedido}")


def anular_linea_pedido_digital(id_pedido, producto):
    """YA NO SE USA (se conserva para que nadie reciba un error de 'funcion inexistente').

    Antes cambiaba el ESTADO de UNA linea a 'ANULADO'. Decision del usuario
    (Regla #30, 26-sep-2026): una comanda es una unidad; las lineas solo se
    quitan en la conversacion ANTES de "Pedido confirmado". Despues, la
    administracion anula la comanda completa con anular_pedido_digital().
    Responde SIEMPRE con un rechazo y no toca el Excel.
    """
    raise AnulacionNoPermitida(
        f"No se anulan lineas sueltas ({producto} en {id_pedido}). Las lineas solo se quitan antes de "
        f"'Pedido confirmado'; despues, la administracion anula la comanda completa con "
        f"anular_pedido_digital(). No se escribio nada."
    )
