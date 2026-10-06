"""Prueba del SaaS de demostracion SOBRE UNA COPIA TEMPORAL (no toca Drive ni la copia local).

Que revisa:
  A) Los candados que protegen el Excel maestro (config.py y datos.py).
  B) Ingreso: sin sesion no se ve nada; clave mala no entra; formulario sin ficha se rechaza.
  C) Permisos: despacho NO puede anular (lo rechaza el servidor, #54) y el Excel no cambia.
  D) "Marcar entregado": escribe la hora en ENTREGADO, ESTADO sigue CONFIRMADO, evento
     ENTREGA_PEDIDO con origen "SaaS (despacho)"; la segunda vez no escribe nada.
  E) "Anular pedido" (administracion): fila -A en negativo, evento ANULACION_PEDIDO,
     'Flujo caja diario' baja en el monto del pedido; se rechazan: motivo vacio,
     pedido pagado por SINPE Movil, pedido ya entregado y pedido ya anulado.
  G) Ventas despacho: comanda de varios productos con TOTAL A COBRAR antes de confirmar (#70),
     una fila por producto en Tabla6, un evento, el Excel rebaja el inventario, envio repetido,
     total que cambio, rechazos del robot y anulacion por el despachador.
  H) Descuento de dia de pago y feriado (#62, 'Interface digital' B5): solo administracion,
     de 0 % a 10 %, evento con valor anterior y nuevo.
  I) Ingreso de Inventario: solo administracion, costo congelado de la TCM, sugerencias de
     nombres, cafe Ca01 y produccion propia rechazados.
  J) Control de inventario (Grupo 2, #63): lo ven los dos usuarios; la existencia minima
     ('Control de Inventario' columna G) la cambia solo administracion, de 0 en adelante con
     hasta 2 decimales; el estado lo recalcula el Excel; evento con valor anterior y nuevo.
  K) Hornadas (Grupo 2): despacho registra y anula; sin materia prima no hay hornada (Regla #34);
     el Excel suma la produccion y rebaja la materia prima; al anular todo vuelve a su valor.
  L) NODO INTEGRAL (Grupo 3, #64): lo ven los dos usuarios; las celdas manuales (C3 tarifa por km,
     C4 meta, C8 accion Meta Ads, C22 SINPE) las cambia solo administracion; C14 y C15 solo lectura;
     'Interface digital' B4 lee la tarifa nueva; evento con valor anterior y nuevo.
     "Oferta hoy" (C16 'Activacion de Promocion', #84): la escribe el robot; la pantalla solo la muestra.
     || 4-oct-2026 (noche): se quito el boton del SaaS (decision del usuario); nadie la actualiza a mano.
  M) Tablas cruce maestras (Grupo 3): solo lectura para los dos usuarios, sin formularios.
  N) Tabla combos dinamicos (Grupo 3, #65): la ven los dos usuarios; la configuracion de cada combo
     (detonador, operador, valor limite, prioridad, mensaje y descuento) la cambia solo administracion;
     el Excel recalcula el estado y el precio; evento con valor anterior y nuevo; el robot revisa C16.
  O) Composicion de los combos (Grupo 3, #85): la ven los dos usuarios; la cambia solo administracion
     (un producto de produccion propia y uno de reventa, cantidad entera de 1 a 10); el robot congela
     lo ya rebajado del inventario y reescribe el precio, el detalle y el mensaje; evento con lo
     anterior y lo nuevo.
  P) Hojas de costo (Grupo 4): 'costo pan baguette' y 'costo cupcakes' las ven los dos usuarios (elementos
     del costo, costo unitario, precio de referencia y utilidad); el precio de referencia (G28 / G27, Regla #9)
     lo cambia solo administracion; bajo el costo unitario lo bloquea el robot; la tabla cruce maestra y los
     combos leen el precio nuevo; evento con valor anterior y nuevo. Nota de C8 en NODO INTEGRAL (#96).
     El cambio se prueba con el producto que tenga el costo unitario por debajo de su precio (la seccion K
     deja una hornada de 10 Baguette, y con tan pocas unidades el costo unitario supera al precio).
  F) La copia local datos_prueba/maestro_prueba.xlsx quedo IGUAL (misma huella).

Necesita: LibreOffice (igual que las otras pruebas) y Flask.
Escribe 13 veces en la copia temporal (unos 10 s cada una en la PC): tarda cerca de 3 minutos.
|| 4-oct-2026: se agregaron J y K (Grupo 2), con 5 escrituras mas; antes tardaba cerca de 2 minutos.
|| 4-oct-2026 (noche): se agregaron L y M (Grupo 3a), con 4 escrituras mas: ahora son 17 y tarda cerca de 3.5 minutos.
|| 5-oct-2026: se agrego N (combos, Grupo 3b), con 4 escrituras mas: ahora son 21 y tarda cerca de 4.5 minutos.
|| 5-oct-2026 (tarde): se agrego O (composicion de los combos), con 3 escrituras mas: ahora son 24 y tarda cerca de 5 minutos.
|| 5-oct-2026 (noche): se agrego P (hojas de costo, Grupo 4), con 2 escrituras mas: ahora son 26 y tarda cerca de 6.5 minutos.

Uso (desde la carpeta ROBOT PANADERO):
    py -m saas_demo.probar_saas_demo
"""
import hashlib
import html
import re
import sys

from .ensayo import COPIA_LOCAL, preparar_copia_temporal

COPIA_TEMPORAL = preparar_copia_temporal()        # ANTES de importar el SaaS: modo ensayo, copia temporal

try:
    from . import config, datos                                  # noqa: E402
    from .app import app                                         # noqa: E402
except ModuleNotFoundError as error:
    if error.name == "flask":
        sys.exit("Falta la libreria Flask. En la Terminal (carpeta ROBOT PANADERO):\n"
                 "    py -m pip install -r whatsapp_bot/requirements.txt")
    raise
from conexion_drive import config as config_robot                # noqa: E402
from conexion_drive import leer_excel                            # noqa: E402
from conexion_drive.recalculo import buscar_errores             # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla      # noqa: E402
from microservicios.grupo1_captura_operativa import descuento_general as dg   # noqa: E402
from microservicios.grupo1_captura_operativa import pedidos_digitales as pd   # noqa: E402
from microservicios.grupo1_captura_operativa.catalogo import Catalogo          # noqa: E402
from microservicios.grupo2_inventario import existencia_minima as em          # noqa: E402
from microservicios.grupo2_inventario import hornadas as hr                   # noqa: E402
from microservicios.grupo3_integraciones import activacion_promocion as ap    # noqa: E402
from microservicios.grupo3_integraciones import combos_composicion as co      # noqa: E402
from microservicios.grupo3_integraciones import combos_config as cc           # noqa: E402
from microservicios.grupo3_integraciones import nodo_manual as nm             # noqa: E402
from microservicios.grupo4_costeo import costeo as cs                         # noqa: E402

errores = []


def paso(titulo):
    print(f"\n--- {titulo} ---")


def revisar(condicion, texto):
    if condicion:
        print(f"  OK     {texto}")
    else:
        print(f"  ERROR  {texto}")
        errores.append(texto)


def huella(ruta):
    return hashlib.md5(ruta.read_bytes()).hexdigest()


def fila_interface(id_pedido):
    wb = leer_excel()
    ws, tabla = obtener_tabla(wb, pd.HOJA_INTERFACE, pd.TABLA_INTERFACE)
    return next((f for f in leer_tabla(ws, tabla) if str(f.get("ID_PEDIDO")).strip() == id_pedido), None)


def ultimo_evento():
    wb = leer_excel()
    ws, tabla = obtener_tabla(wb, "RegistroEventos", "tblRegistroEventos")
    return leer_tabla(ws, tabla)[-1]


def suma_flujo(columna):
    """Suma de una columna de 'Flujo caja diario' (3 = Ventas digitales, 11 = Envios cobrados)."""
    ws = leer_excel()["Flujo caja diario"]
    return round(sum(ws.cell(f, columna).value or 0 for f in range(5, ws.max_row + 1)
                     if isinstance(ws.cell(f, columna).value, (int, float))), 2)


def unico_de(pagina):
    """El numero de envio unico que la pantalla trae escondido en el formulario."""
    return re.search(r'name="unico" value="([^"]+)"', pagina.get_data(as_text=True)).group(1)


def inventario(producto, columna):
    """Un valor de 'Control de Inventario' (lo calcula el Excel con sus formulas)."""
    wb = leer_excel()
    ws, tabla = obtener_tabla(wb, "Control de Inventario", "TablaControlInventario")
    fila = next(f for f in leer_tabla(ws, tabla) if str(f.get("Producto Maestro")).strip() == producto)
    return round(float(fila.get(columna) or 0), 2)


def ultima_fila(hoja, nombre_tabla, columna_clave):
    wb = leer_excel()
    ws, tabla = obtener_tabla(wb, hoja, nombre_tabla)
    return [f for f in leer_tabla(ws, tabla) if f.get(columna_clave) not in (None, "")][-1]


def configuracion_rechazada(entorno, fragmento):
    try:
        config.resolver(entorno)
    except config.ConfiguracionInvalida as error:
        return fragmento.lower() in str(error).lower()
    return False


def ficha_de(cliente):
    cliente.get("/entrar")
    with cliente.session_transaction() as sesion:
        return sesion["ficha"]


def ingresar(cliente, rol):
    cuenta = next(c for c in config.USUARIOS.values() if c["rol"] == rol)
    ficha = ficha_de(cliente)
    respuesta = cliente.post("/entrar", data={"usuario": cuenta["usuario"], "clave": cuenta["clave"], "ficha": ficha})
    return respuesta, ficha


def main():
    huella_local_antes = huella(COPIA_LOCAL)
    print(f"Copia temporal de la prueba: {COPIA_TEMPORAL}")

    # ------------------------------------------------------------------ A
    paso("A) Candados que protegen el Excel maestro")
    claves = {"SAAS_CLAVE_DESPACHO": "clave-despacho-1", "SAAS_CLAVE_ADMIN": "clave-admin-22"}
    id_a, id_b = "1AAAAaaaaBBBBbbbbCCCCcccc0001", "1ZZZZzzzzYYYYyyyyXXXXxxxx0002"
    revisar(configuracion_rechazada({**claves}, "Falta SAAS_EXCEL_FILE_ID"),
            "sin SAAS_EXCEL_FILE_ID el SaaS no arranca")
    revisar(configuracion_rechazada({**claves, "SAAS_EXCEL_FILE_ID": id_a, "EXCEL_MAESTRO_FILE_ID": id_a}, "MISMO archivo"),
            "si el ID de la copia es el del maestro, no arranca")
    enlace = f"https://docs.google.com/spreadsheets/d/{id_a}/edit?usp=drive_link"
    revisar(configuracion_rechazada({**claves, "SAAS_EXCEL_FILE_ID": enlace, "EXCEL_MAESTRO_FILE_ID": id_a}, "MISMO archivo"),
            "tambien lo detecta si se pega el ENLACE en vez del ID")
    revisar(configuracion_rechazada({"SAAS_EXCEL_FILE_ID": id_b}, "Falta SAAS_CLAVE_DESPACHO"),
            "sin claves no arranca")
    revisar(configuracion_rechazada({"SAAS_EXCEL_FILE_ID": id_b, "SAAS_CLAVE_DESPACHO": "corta", "SAAS_CLAVE_ADMIN": "clave-admin-22"},
                                    "al menos"), "una clave de menos de 8 caracteres se rechaza")
    revisar(configuracion_rechazada({"SAAS_EXCEL_FILE_ID": id_b, "SAAS_CLAVE_DESPACHO": "misma-clave-1", "SAAS_CLAVE_ADMIN": "misma-clave-1"},
                                    "distintas"), "las dos claves deben ser distintas")
    buena = config.resolver({**claves, "SAAS_EXCEL_FILE_ID": enlace, "EXCEL_MAESTRO_FILE_ID": id_b})
    revisar(buena["id_demo"] == id_a and not buena["ensayo"], "configuracion correcta: toma el ID de la copia desde el enlace")
    revisar(not datos.nombre_es_copia("Estudio_Financiero_Panadero_CON_ROUND_v68_DRIVE_CALCULADO.xlsx"),
            "el nombre del Excel maestro NO pasa por copia (no contiene DEMO)")
    revisar(datos.nombre_es_copia("DEMO_Estudio_Financiero_Panadero.xlsx"), "un nombre con DEMO si pasa")

    # Simular el modo Drive con un archivo que NO es la copia: ni lee ni escribe
    originales = (config.ENSAYO, config.ID_DEMO, config_robot.ORIGEN_EXCEL, config_robot.EXCEL_MAESTRO_FILE_ID,
                  datos.obtener_metadatos, pd.marcar_entregado)
    llamadas = []
    try:
        config.ENSAYO, config.ID_DEMO = False, id_a
        config_robot.ORIGEN_EXCEL, config_robot.EXCEL_MAESTRO_FILE_ID = "drive", id_a
        datos.obtener_metadatos = lambda: {"name": "Estudio_Financiero_Panadero.xlsx", "version": "1"}
        pd.marcar_entregado = lambda *a, **k: llamadas.append(a)
        try:
            datos.entregar("WA-000", "despacho")
            rechazo = False
        except datos.CopiaNoVerificada:
            rechazo = True
        revisar(rechazo and not llamadas, "si el archivo de Drive no se llama DEMO, el boton NO llega a escribir")
        config_robot.EXCEL_MAESTRO_FILE_ID = id_b
        try:
            datos.verificar_copia(reciente=True)
            rechazo = False
        except datos.CopiaNoVerificada:
            rechazo = True
        revisar(rechazo, "si el robot apuntara a otro ID distinto del de la copia, se detiene")
    finally:
        (config.ENSAYO, config.ID_DEMO, config_robot.ORIGEN_EXCEL, config_robot.EXCEL_MAESTRO_FILE_ID,
         datos.obtener_metadatos, pd.marcar_entregado) = originales
        datos.olvidar_cache()

    # ------------------------------------------------------------------ B
    paso("B) Ingreso")
    despacho = app.test_client()
    revisar(despacho.get("/salud").get_json().get("modo") == "ensayo", "/salud responde y dice modo ensayo")
    respuesta = despacho.get("/")
    revisar(respuesta.status_code == 302 and "/entrar" in respuesta.headers["Location"], "sin sesion, la lista manda a /entrar")
    ficha = ficha_de(despacho)
    respuesta = despacho.post("/entrar", data={"usuario": "despacho", "clave": "no-es-la-clave", "ficha": ficha})
    revisar(respuesta.status_code == 401, "clave equivocada: no entra")
    respuesta = despacho.post("/entrar", data={"usuario": "despacho", "clave": config.CLAVE_ENSAYO})
    revisar(respuesta.status_code == 400, "formulario sin ficha: rechazado (400)")
    respuesta, ficha_despacho = ingresar(despacho, config.ROL_DESPACHO)
    revisar(respuesta.status_code == 302, "despacho ingresa")
    pagina = despacho.get("/?ver=todos")
    libro = datos.leer_pedidos()
    revisar(pagina.status_code == 200 and all(p["id"].encode() in pagina.data for p in libro["pedidos"]),
            f"la lista muestra los {len(libro['pedidos'])} pedidos del Excel")

    pendientes = [p for p in libro["pedidos"] if p["situacion"] == datos.PENDIENTE]
    efectivo = [p for p in pendientes if p["puede_anular"]]
    sinpe = [p for p in pendientes if not p["puede_anular"]]
    if len(efectivo) < 2 or not sinpe:
        print("\nLa copia local no tiene suficientes pedidos por entregar (2 en Efectivo y 1 por SINPE). No se puede seguir.")
        return 1
    para_entregar, para_anular, pagado_sinpe = efectivo[0]["id"], efectivo[1]["id"], sinpe[0]["id"]
    print(f"  (pedidos de la prueba: entregar {para_entregar}, anular {para_anular}, SINPE {pagado_sinpe})")

    # ------------------------------------------------------------------ C
    paso("C) Despacho NO puede anular")
    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = despacho.post(f"/pedidos/{para_anular}/anular", data={"motivo": "prueba", "ficha": ficha_despacho})
    revisar(respuesta.status_code == 403, "el servidor responde 403 a despacho")
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "el Excel no cambio")
    pagina = despacho.get(f"/pedidos/{para_anular}")
    revisar(b"Anular pedido</button>" not in pagina.data, "despacho no ve el boton de anular")

    # ------------------------------------------------------------------ D
    paso("D) Marcar entregado (despacho)")
    respuesta = despacho.post(f"/pedidos/{para_entregar}/entregar", data={"ficha": ficha_despacho})
    revisar(respuesta.status_code == 302, "el boton responde y vuelve al pedido")
    fila = fila_interface(para_entregar)
    revisar(fila.get(pd.COL_ENTREGADO) not in (None, ""), f"{para_entregar}: ENTREGADO tiene fecha y hora ({fila.get(pd.COL_ENTREGADO)})")
    revisar(fila.get("ESTADO") == pd.CONFIRMADO, "ESTADO sigue CONFIRMADO (Regla #35)")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == pd.TIPO_EVENTO_ENTREGA and evento.get("ORIGEN") == "SaaS (despacho)"
            and evento.get("ID_ENTIDAD") == para_entregar, f"evento {evento.get('ID_EVENTO')}: ENTREGA_PEDIDO, origen SaaS (despacho)")
    pagina = despacho.get(f"/pedidos/{para_entregar}")
    revisar("quedó entregado".encode() in pagina.data and b"Recibo" in pagina.data, "la pantalla muestra el recibo y el tiempo")
    huella_antes = huella(COPIA_TEMPORAL)
    despacho.post(f"/pedidos/{para_entregar}/entregar", data={"ficha": ficha_despacho})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "segunda vez: ya estaba entregado, no se escribe nada")

    # ------------------------------------------------------------------ E
    paso("E) Anular pedido (administracion)")
    admin = app.test_client()
    respuesta, ficha_admin = ingresar(admin, config.ROL_ADMIN)
    revisar(respuesta.status_code == 302, "administracion ingresa")
    huella_antes = huella(COPIA_TEMPORAL)
    for id_pedido, motivo, texto in (
        (para_anular, "   ", "motivo vacio: rechazado"),
        (pagado_sinpe, "prueba", f"{pagado_sinpe} pagado por SINPE Movil: rechazado (Regla #30)"),
        (para_entregar, "prueba", f"{para_entregar} ya entregado: rechazado"),
    ):
        admin.post(f"/pedidos/{id_pedido}/anular", data={"motivo": motivo, "ficha": ficha_admin})
        pagina = admin.get(f"/pedidos/{id_pedido}")
        revisar("El robot no lo permitió".encode() in pagina.data and huella(COPIA_TEMPORAL) == huella_antes, texto)

    original = fila_interface(para_anular)
    ventas_antes, envios_antes = suma_flujo(3), suma_flujo(11)
    respuesta = admin.post(f"/pedidos/{para_anular}/anular",
                           data={"motivo": "El cliente no estaba en la direccion", "ficha": ficha_admin})
    revisar(respuesta.status_code == 302, "el boton responde y vuelve al pedido")
    anulacion = fila_interface(para_anular + pd.SUFIJO_ANULACION)
    revisar(anulacion is not None and anulacion.get("TOTAL_A_COBRAR") == -original.get("TOTAL_A_COBRAR"),
            f"fila {para_anular}-A con el total en negativo ({anulacion.get('TOTAL_A_COBRAR') if anulacion else None})")
    revisar(fila_interface(para_anular).get("TOTAL_A_COBRAR") == original.get("TOTAL_A_COBRAR"),
            "el pedido original no se modifico")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == pd.TIPO_EVENTO_ANULACION and evento.get("ORIGEN") == "SaaS (administracion)",
            f"evento {evento.get('ID_EVENTO')}: ANULACION_PEDIDO, origen SaaS (administracion)")
    comanda = round((original.get("SUBTOTAL_BRUTO") or 0) - (original.get("DESCUENTO_TOTAL") or 0), 2)
    envio = round(original.get("ENVIO_MAPS") or 0, 2)
    revisar(round(ventas_antes - suma_flujo(3), 2) == comanda, f"'Flujo caja diario' ventas digitales bajo {comanda:,.2f}")
    revisar(round(envios_antes - suma_flujo(11), 2) == envio, f"'Flujo caja diario' envios cobrados bajo {envio:,.2f}")
    huella_antes = huella(COPIA_TEMPORAL)
    admin.post(f"/pedidos/{para_anular}/anular", data={"motivo": "otra vez", "ficha": ficha_admin})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "anular dos veces: rechazado, no se escribe nada")
    admin.post(f"/pedidos/{para_anular}/entregar", data={"ficha": ficha_admin})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "entregar un pedido anulado: rechazado, no se escribe nada")
    pagina = admin.get("/?ver=anulados")
    revisar(para_anular.encode() in pagina.data, "el pedido aparece en la lista de Anulados")
    respuesta = admin.post("/salir", data={"ficha": ficha_admin})
    revisar(admin.get("/").status_code == 302, "despues de Salir ya no se ve la lista")

    # ------------------------------------------------------------------ G
    paso("G) Ventas despacho: comanda con total a cobrar (#70), despacho registra y anula")
    pagina = despacho.get("/ventas-despacho")
    revisar(pagina.status_code == 200 and "La comanda está vacía".encode() in pagina.data, "la pantalla abre con la comanda vacia")
    catalogo = Catalogo(leer_excel())
    precios = {c: round(float(p), 2) for c, n, p in catalogo.lista_para_mostrar()}
    leche = next(n for c, n, p in catalogo.lista_para_mostrar() if c == "L01")
    existencia_antes, leche_antes = inventario("Baguette", "Existencia Actual"), inventario(leche, "Existencia Actual")
    huella_antes = huella(COPIA_TEMPORAL)

    def agregar(producto, cantidad):
        return despacho.post("/ventas-despacho/agregar", data={"producto": producto, "cantidad": cantidad, "ficha": ficha_despacho})

    inactivo = next((c["codigo"] for c in catalogo.combos if c["codigo"].startswith("co") and not c["activo"]), None)
    for producto, cantidad in (("P01", "0"), ("XX99", "1"), (inactivo, "1")):
        if producto:
            agregar(producto, cantidad)
    pagina = despacho.get("/ventas-despacho")
    revisar("La comanda está vacía".encode() in pagina.data,
            f"cantidad 0, producto inexistente y combo {inactivo} inactivo: el robot no los deja entrar a la comanda")
    agregar("P01", "2"), agregar("L01", "1"), agregar("P01", "1")
    esperado = round(3 * precios["P01"] + precios["L01"], 2)
    pagina = despacho.get("/ventas-despacho")
    revisar(f"₡{esperado:,.2f}".encode() in pagina.data and b"Total a cobrar" in pagina.data,
            f"3 Baguette + 1 leche: la pantalla muestra el total a cobrar {esperado:,.2f} antes de confirmar")
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "armar la comanda no escribe nada en el Excel")
    despacho.post("/ventas-despacho/registrar", data={"unico": unico_de(pagina), "total": str(esperado + 100), "ficha": ficha_despacho})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "si el total ya no es el que se mostro: rechazado, no se escribe nada")
    pagina = despacho.get("/ventas-despacho")
    unico = unico_de(pagina)
    despacho.post("/ventas-despacho/registrar", data={"unico": unico, "total": str(esperado), "ficha": ficha_despacho})
    ws_v, tabla_v = obtener_tabla(leer_excel(), "Ventas despacho", "Tabla6")
    filas_v = [f for f in leer_tabla(ws_v, tabla_v) if f.get("Producto") not in (None, "")][-2:]
    revisar([(f.get("Producto"), f.get("Cantidad")) for f in filas_v] == [("Baguette", 3), (leche, 1)]
            and filas_v[0].get("Hora") == filas_v[1].get("Hora") and all(f.get("Estado de orden") == "ORDEN CORRECTA" for f in filas_v),
            f"filas {filas_v[0]['_fila']} y {filas_v[1]['_fila']} de Tabla6: una por producto, misma fecha y hora")
    revisar(filas_v[0].get("Precio") == precios["P01"] and round(sum(f.get("Total") for f in filas_v), 2) == esperado,
            f"precios de la tabla cruce maestra; los totales por formula suman {esperado:,.2f}")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == "VENTA_DESPACHO" and evento.get("ORIGEN") == "SaaS (despacho)"
            and evento.get("VALOR_REFERENCIA") == esperado, f"un evento VENTA_DESPACHO por la comanda, con el total {esperado:,.2f}")
    # round(): 8.55 - 1 da 7.550000000000001 en la computadora; se compara a 2 decimales.
    revisar(inventario("Baguette", "Existencia Actual") == round(existencia_antes - 3, 2)
            and inventario(leche, "Existencia Actual") == round(leche_antes - 1, 2),
            f"el Excel rebajo el inventario: Baguette {existencia_antes:g} -> {existencia_antes - 3:g}, leche {leche_antes:g} -> {leche_antes - 1:g}")
    revisar("La comanda está vacía".encode() in despacho.get("/ventas-despacho").data, "despues de registrar, la comanda queda vacia")
    huella_antes = huella(COPIA_TEMPORAL)
    agregar("P01", "3"), agregar("L01", "1")
    despacho.post("/ventas-despacho/registrar", data={"unico": unico, "total": str(esperado), "ficha": ficha_despacho})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "el mismo formulario enviado dos veces no registra dos ventas")
    despacho.post("/ventas-despacho/quitar", data={"producto": "L01", "ficha": ficha_despacho})
    revisar(f"₡{3 * precios['P01']:,.2f}".encode() in despacho.get("/ventas-despacho").data, "quitar una linea: el total a cobrar se recalcula")
    despacho.post("/ventas-despacho/quitar", data={"producto": "*", "ficha": ficha_despacho})
    despacho.post("/ventas-despacho/registrar", data={"unico": "prueba-comanda-vacia", "total": "0", "ficha": ficha_despacho})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "comanda vacia: no se registra nada")
    for fila in filas_v:
        despacho.post("/ventas-despacho/anular", data={"fila": str(fila["_fila"]), "ficha": ficha_despacho})
    ws_v, tabla_v = obtener_tabla(leer_excel(), "Ventas despacho", "Tabla6")
    anuladas = [f for f in leer_tabla(ws_v, tabla_v) if f["_fila"] in (filas_v[0]["_fila"], filas_v[1]["_fila"])]
    revisar(all(f.get("Estado de orden") == "ORDEN ANULADA" for f in anuladas),
            "anular: solo cambia el estado a ORDEN ANULADA, producto por producto (las ventas no se borran)")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == "ANULACION" and evento.get("ORIGEN") == "SaaS (despacho)",
            f"evento {evento.get('ID_EVENTO')}: ANULACION, origen SaaS (despacho) (antes decia Robot Python)")
    revisar(inventario("Baguette", "Existencia Actual") == existencia_antes and inventario(leche, "Existencia Actual") == leche_antes,
            "el inventario volvio a su valor")
    huella_antes = huella(COPIA_TEMPORAL)
    despacho.post("/ventas-despacho/anular", data={"fila": str(filas_v[0]["_fila"]), "ficha": ficha_despacho})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "anular dos veces: no se escribe nada")

    # ------------------------------------------------------------------ H
    paso("H) Descuento de dia de pago y feriado (#62, 'Interface digital' B5)")
    revisar(despacho.get("/descuento-general").status_code == 403, "despacho no puede entrar (403)")
    admin = app.test_client()
    respuesta, ficha_admin = ingresar(admin, config.ROL_ADMIN)
    pagina = admin.get("/descuento-general")
    inicial = dg.leer_descuento_general()
    revisar(pagina.status_code == 200 and f"Hoy está en {inicial:g} %".encode() in pagina.data,
            f"administracion ve el valor actual ({inicial:g} %)")
    huella_antes = huella(COPIA_TEMPORAL)
    for valor in ("11", "-1", "abc"):
        admin.post("/descuento-general", data={"porcentaje": valor, "ficha": ficha_admin})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "11, -1 y 'abc': rechazados, no se escribe nada (maximo 10 %)")
    nuevo = 10 if inicial != 10 else 5
    admin.post("/descuento-general", data={"porcentaje": str(nuevo), "ficha": ficha_admin})
    revisar(leer_excel()["Interface digital"]["B5"].value == nuevo / 100, f"B5 quedo en {nuevo / 100} ({nuevo} %)")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == dg.TIPO_EVENTO and evento.get("ORIGEN") == "SaaS (administracion)"
            and "->" in str(evento.get("DETALLE")), f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    revisar(Catalogo(leer_excel()).reglas["pct_general"] == nuevo / 100, "el robot (catalogo.py) ya lee el descuento nuevo")
    huella_antes = huella(COPIA_TEMPORAL)
    admin.post("/descuento-general", data={"porcentaje": str(nuevo), "ficha": ficha_admin})
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "el mismo valor otra vez: SIN_CAMBIO, no se escribe nada")
    revisar(dg.normalizar_porcentaje("7,5 %") == 7.5, "la funcion acepta '7,5 %' como 7.5 %")
    admin.post("/descuento-general", data={"porcentaje": f"{inicial:g}", "ficha": ficha_admin})
    revisar(dg.leer_descuento_general() == inicial, f"el descuento volvio a {inicial:g} %")

    # ------------------------------------------------------------------ I
    paso("I) Ingreso de Inventario (administracion)")
    revisar(despacho.get("/ingreso-inventario").status_code == 403, "despacho no puede entrar (403)")
    pagina = admin.get("/ingreso-inventario")
    revisar(pagina.status_code == 200 and b"Ha01" in pagina.data and b'value="Ca01"' not in pagina.data
            and b'value="P01"' not in pagina.data, "la lista ofrece Ha01 y no ofrece Ca01 ni produccion propia")
    wb = leer_excel()
    harina = next(f["Materia Prima"] for f in leer_tabla(*obtener_tabla(wb, "TABLA CRUCE MAESTRA MP", "TablaCruceMaestraMP"))
                  if str(f.get("Codigo")).strip() == "Ha01")
    costo_harina = next(f["Costo Unitario (Scraping)"] for f in leer_tabla(*obtener_tabla(wb, "TABLA CRUCE MAESTRA MP", "TablaCruceMaestraMP"))
                        if str(f.get("Codigo")).strip() == "Ha01")
    entrada_antes = inventario(harina, "Entrada Compra")

    def factura(lineas, proveedor="Solera", codigo="F-SAAS-01"):
        datos_formulario = {"fecha": "2026-10-03", "factura": codigo, "proveedor": proveedor, "ficha": ficha_admin,
                            "unico": unico_de(admin.get("/ingreso-inventario")),
                            "producto": [l[0] for l in lineas], "cantidad": [l[1] for l in lineas]}
        return admin.post("/ingreso-inventario", data=datos_formulario)

    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = factura([("HARINA FLORES 5 KG", "2")])
    revisar(huella(COPIA_TEMPORAL) == huella_antes and "Quiso decir".encode() in respuesta.data,
            "nombre mal escrito: rechazado y el robot sugiere el nombre correcto")
    respuesta = factura([("Ca01", "1")])
    revisar(huella(COPIA_TEMPORAL) == huella_antes and b"Cf01" in respuesta.data,
            "cafe con el codigo Ca01: rechazado, indica comprar con Cf01 (Regla #32)")
    respuesta = factura([("P01", "50")])
    revisar(huella(COPIA_TEMPORAL) == huella_antes and b"Hornadas" in respuesta.data,
            "produccion propia (P01): rechazada, se registra en Hornadas")
    respuesta = factura([("Ha01", "2")], proveedor="")
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "sin proveedor: rechazada, no se escribe nada")
    respuesta = factura([("Ha01", "2"), ("L01", "3")])
    revisar(respuesta.status_code == 302, "factura F-SAAS-01 con 2 lineas: aceptada")
    ws_i, tabla_i = obtener_tabla(leer_excel(), "Ingreso de Inventario", "TablaIngresos")
    nuevas = [f for f in leer_tabla(ws_i, tabla_i) if f.get("Factura") == "F-SAAS-01"]
    revisar(len(nuevas) == 2 and nuevas[0].get("Compra externa") == harina and nuevas[0].get("Cantidad Compras") == 2
            and nuevas[0].get("Proveedor") == "Solera", "2 filas nuevas en TablaIngresos con el nombre exacto y el proveedor")
    revisar(nuevas[0].get("Precio unitario") == costo_harina and round(nuevas[0].get("Total Compra Externa"), 2) == round(2 * costo_harina, 2),
            f"costo congelado de la TCM ({costo_harina:,.2f}); nadie digito el precio (Regla #33)")
    revisar(ultimo_evento().get("TIPO_EVENTO") == "INGRESO_INVENTARIO" and ultimo_evento().get("ID_ENTIDAD") == "F-SAAS-01",
            "evento INGRESO_INVENTARIO de la factura")
    revisar(ultimo_evento().get("ORIGEN") == "SaaS (administracion)",
            "el evento dice origen SaaS (administracion) (antes decia Robot Python)")
    revisar(inventario(harina, "Entrada Compra") == round(entrada_antes + 2, 2),
            f"el Excel sumo la compra al inventario de harina: {entrada_antes:g} -> {entrada_antes + 2:g}")
    huella_antes = huella(COPIA_TEMPORAL)
    factura([("Ha01", "2"), ("L01", "3")])
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "la misma factura otra vez: duplicada, no se escribe nada")

    # ------------------------------------------------------------------ J
    paso("J) Control de inventario: lectura para todos; existencia minima solo administracion (#63)")
    filas_inv = em.leer_existencias()
    pagina = despacho.get("/control-inventario")
    texto = html.unescape(pagina.get_data(as_text=True))
    revisar(pagina.status_code == 200 and all(f["producto"] in texto for f in filas_inv),
            f"despacho ve los {len(filas_inv)} productos de 'Control de Inventario'")
    revisar(b"Guardar existencia" not in pagina.data, "despacho no ve el formulario de la existencia minima")
    elegido = next((f for f in filas_inv if f["estado"] == "REABASTECER" and f["existencia"] >= 0.1), filas_inv[0])

    def minimo(cliente, ficha, producto, valor):
        return cliente.post("/control-inventario/minimo", data={"producto": producto, "minimo": valor, "ficha": ficha})

    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = minimo(despacho, ficha_despacho, elegido["producto"], "1")
    revisar(respuesta.status_code == 403 and huella(COPIA_TEMPORAL) == huella_antes,
            "despacho no puede cambiarla: el servidor responde 403 y el Excel no cambia")
    revisar(b"Guardar existencia" in admin.get("/control-inventario").data, "administracion si ve el formulario")
    for producto, valor in ((elegido["producto"], "-1"), (elegido["producto"], "abc"), ("PRODUCTO QUE NO EXISTE", "5")):
        minimo(admin, ficha_admin, producto, valor)
    pagina = admin.get("/control-inventario")
    revisar("El robot no lo permitió".encode() in pagina.data and huella(COPIA_TEMPORAL) == huella_antes,
            "-1, 'abc' y un producto que no esta en la hoja: rechazados, no se escribe nada")
    nuevo_minimo = max(round(elegido["existencia"] / 2, 2), 0.01)
    minimo(admin, ficha_admin, elegido["producto"], f"{nuevo_minimo:g}".replace(".", ","))
    fila_inv = next(f for f in em.leer_existencias() if f["producto"] == elegido["producto"])
    estado_esperado = "REABASTECER" if round(elegido["existencia"], 2) < nuevo_minimo else "OK"
    revisar(fila_inv["minimo"] == nuevo_minimo, f"{elegido['producto']}: minimo {elegido['minimo']:g} -> {nuevo_minimo:g} (con decimales)")
    revisar(fila_inv["estado"] == estado_esperado and round(fila_inv["existencia"], 2) == round(elegido["existencia"], 2),
            f"el Excel recalculo el estado ({elegido['estado']} -> {fila_inv['estado']}); la existencia no cambio")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == em.TIPO_EVENTO and evento.get("ORIGEN") == "SaaS (administracion)"
            and "->" in str(evento.get("DETALLE")), f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    revisar(b"Guardado en el Excel en" in admin.get("/control-inventario").data, "la pantalla muestra el aviso con el tiempo")
    huella_antes = huella(COPIA_TEMPORAL)
    minimo(admin, ficha_admin, elegido["producto"], f"{nuevo_minimo:g}")
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "el mismo valor otra vez: SIN_CAMBIO, no se escribe nada")
    minimo(admin, ficha_admin, elegido["producto"], f"{elegido['minimo']:g}")
    revisar([(f["producto"], f["minimo"], f["estado"]) for f in em.leer_existencias()]
            == [(f["producto"], f["minimo"], f["estado"]) for f in filas_inv], "los minimos y los estados volvieron a su valor")

    # ------------------------------------------------------------------ K
    paso("K) Hornadas: despacho registra y anula; sin materia prima no hay hornada (Regla #34)")
    pagina = despacho.get("/hornadas")
    revisar(pagina.status_code == 200 and b'value="P01"' in pagina.data and b'value="Pa01"' in pagina.data,
            "la pantalla ofrece los dos productos de produccion propia (P01 y Pa01)")

    def hornear(producto, cantidad, unico=None):
        unico = unico or unico_de(despacho.get("/hornadas"))
        return despacho.post("/hornadas/registrar", data={"producto": producto, "cantidad": cantidad, "unico": unico,
                                                         "ficha": ficha_despacho})

    def consumo(materia):
        """'Consumo en Produccion' de una materia prima, sin redondear (lo calcula el Excel)."""
        ws_c, tabla_c = obtener_tabla(leer_excel(), "Control de Inventario", "TablaControlInventario")
        fila = next(f for f in leer_tabla(ws_c, tabla_c) if str(f.get("Producto Maestro")).strip() == materia)
        return float(fila.get("Consumo en Producción") or 0)

    huella_antes = huella(COPIA_TEMPORAL)
    hornear("P01", "9999")
    pagina = despacho.get("/hornadas")
    revisar(huella(COPIA_TEMPORAL) == huella_antes and b"No alcanza la materia prima" in pagina.data,
            "9999 Baguette: el robot la rechaza completa, dice que falta y no escribe nada")
    hornear("L01", "10"), hornear("P01", "0"), hornear("P01", "2.5")
    despacho.get("/hornadas")
    revisar(huella(COPIA_TEMPORAL) == huella_antes,
            "producto que no es de produccion propia, cantidad 0 y cantidad con decimales: no se escribe nada")

    # Primero se compra y despues se hornea: 1 paquete de cada materia prima de la receta de Baguette.
    ws_r, tabla_r = obtener_tabla(leer_excel(), hr.HOJA_CONSUMO, hr.TABLA_RECETA)
    receta = [l for l in leer_tabla(ws_r, tabla_r) if str(l.get("Codigo")).strip() == "P01"]
    materias = [str(l["Materia prima"]).strip() for l in receta]
    respuesta = factura([(m, "1") for m in materias], codigo="F-SAAS-02")
    revisar(respuesta.status_code == 302, f"compra F-SAAS-02: 1 paquete de cada materia prima de la receta ({len(materias)} lineas)")

    unidades = 10
    produccion_antes = inventario("Baguette", "Entrada Producción")
    existencia_antes = inventario("Baguette", "Existencia Actual")
    consumo_antes = {m: consumo(m) for m in materias}
    unico = unico_de(despacho.get("/hornadas"))
    respuesta = hornear("P01", str(unidades), unico)
    revisar(respuesta.status_code == 302, f"hornada de {unidades} Baguette: el boton responde")
    hist = ultima_fila("Historico Diario", "HistoricoHorneadas", "Fecha")
    revisar(hist.get("Producto") == "Baguette" and hist.get("Cantidad producida") == unidades,
            f"'Historico Diario' fila {hist['_fila']}: {unidades} x Baguette")
    revisar(round(hist.get("Venta total") - hist.get("Costo total producción"), 2) == round(hist.get("Utilidad bruta"), 2)
            and round(unidades * hist.get("Precio unitario"), 2) == round(hist.get("Venta total"), 2),
            f"los 9 valores quedaron calculados (costo unitario {hist.get('Costo unitario'):,.2f}, "
            f"precio de referencia {hist.get('Precio unitario'):,.2f})")
    revisar(inventario("Baguette", "Entrada Producción") == round(produccion_antes + unidades, 2)
            and inventario("Baguette", "Existencia Actual") == round(existencia_antes + unidades, 2),
            f"el Excel sumo la produccion: existencia de Baguette {existencia_antes:g} -> {existencia_antes + unidades:g}")
    esperado_mp = {str(l["Materia prima"]).strip(): unidades * l["Cantidad por unidad"] / l["Contenido por paquete"] for l in receta}
    revisar(all(abs(consumo(m) - consumo_antes[m] - esperado_mp[m]) < 0.0002 for m in materias),
            "el Excel rebajo la materia prima de la receta: "
            + ", ".join(f"{m[:10].strip()} {esperado_mp[m]:.3f}" for m in materias) + " paquetes")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == "HORNADA" and evento.get("ORIGEN") == "SaaS (despacho)",
            f"evento {evento.get('ID_EVENTO')}: HORNADA, origen SaaS (despacho), lote {evento.get('ID_ENTIDAD')}")
    pagina = despacho.get("/hornadas")
    revisar(b"Hornada registrada" in pagina.data and b"Guardado en el Excel en" in pagina.data,
            "la pantalla muestra el aviso con el lote, el costo y el tiempo")
    huella_antes = huella(COPIA_TEMPORAL)
    hornear("P01", str(unidades), unico)
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "el mismo formulario enviado dos veces no registra dos hornadas")

    despacho.post("/hornadas/anular", data={"producto": "P01", "ficha": ficha_despacho})
    hist = ultima_fila("Historico Diario", "HistoricoHorneadas", "Fecha")
    revisar(hist.get("Producto") == "Baguette" and hist.get("Cantidad producida") == -unidades,
            f"anular: fila {hist['_fila']} de correccion con cantidad negativa ({hist.get('Cantidad producida')}); la original no se borra")
    revisar(inventario("Baguette", "Existencia Actual") == existencia_antes
            and all(abs(consumo(m) - consumo_antes[m]) < 0.0002 for m in materias),
            "el inventario de Baguette y la materia prima volvieron a su valor")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == "ANULACION_HORNADA" and evento.get("ORIGEN") == "SaaS (despacho)",
            f"evento {evento.get('ID_EVENTO')}: ANULACION_HORNADA, origen SaaS (despacho)")
    huella_antes = huella(COPIA_TEMPORAL)
    despacho.post("/hornadas/anular", data={"producto": "P01", "ficha": ficha_despacho})
    revisar(huella(COPIA_TEMPORAL) == huella_antes and "ya estaba anulada".encode() in despacho.get("/hornadas").data,
            "anular dos veces: ya estaba anulada, no se escribe nada")
    respuesta = despacho.post("/hornadas/anular", data={"producto": "L01", "ficha": ficha_despacho})
    revisar(respuesta.status_code == 404 and huella(COPIA_TEMPORAL) == huella_antes,
            "anular un producto que no es de produccion propia: 404, no se escribe nada")

    # ------------------------------------------------------------------ L
    paso("L) NODO INTEGRAL: lectura para todos; celdas manuales solo administracion (#64)")
    filas_nodo = nm.leer_nodo()
    manuales = {f["campo"]: f for f in filas_nodo if f["clase"] == "manual"}
    pagina = despacho.get("/nodo-integral")
    texto = html.unescape(pagina.get_data(as_text=True))
    revisar(pagina.status_code == 200 and all(f["descripcion"] in texto for f in filas_nodo),
            f"despacho ve las {len(filas_nodo)} filas de 'NODO INTEGRAL'")
    revisar(b"Guardar tarifa" not in pagina.data and b'name="campo"' not in pagina.data,
            "despacho no ve los formularios de las celdas manuales")
    revisar(set(manuales) == set(nm.CAMPOS) and [f["celda"] for f in filas_nodo if f["clase"] == "sin_efecto"] == ["C14", "C15"],
            "manuales: C3, C4, C8 y C22; solo lectura por la Regla #8: C14 y C15")

    def dato(cliente, ficha, campo, valor):
        return cliente.post("/nodo-integral/guardar", data={"campo": campo, "valor": valor, "ficha": ficha})

    tarifa_inicial = manuales["tarifa_km"]["valor"]
    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = dato(despacho, ficha_despacho, "tarifa_km", "600")
    revisar(respuesta.status_code == 403 and huella(COPIA_TEMPORAL) == huella_antes,
            "despacho no puede cambiarlas: el servidor responde 403 y el Excel no cambia")
    pagina = admin.get("/nodo-integral")
    revisar(all(texto_boton.encode() in pagina.data for texto_boton in
                ("Guardar tarifa", "Guardar meta", "Guardar acción", "Guardar número")),
            "administracion ve los cuatro formularios")
    for campo, valor in (("tarifa_km", "0"), ("tarifa_km", "5001"), ("tarifa_km", "500.5"), ("meta_venta", "-1"),
                         ("meta_ads", "Bajar"), ("sinpe", "1234"), ("oferta_feriado", "Combo 1"), ("clima", "Lluvia")):
        dato(admin, ficha_admin, campo, valor)
    pagina = admin.get("/nodo-integral")
    revisar("El robot no lo permitió".encode() in pagina.data and huella(COPIA_TEMPORAL) == huella_antes,
            "tarifa 0, 5001 y 500.5, meta -1, accion 'Bajar', SINPE de 4 digitos, C14 y una fila del robot: "
            "rechazados, no se escribe nada")
    nueva_tarifa = 600 if tarifa_inicial != 600 else 650
    dato(admin, ficha_admin, "tarifa_km", str(nueva_tarifa))
    wb = leer_excel()
    revisar(wb[nm.HOJA]["C3"].value == nueva_tarifa and wb["Interface digital"]["B4"].value == nueva_tarifa,
            f"Tarifa Base KM {tarifa_inicial} -> {nueva_tarifa}: celda C3, y 'Interface digital' B4 (formula) la lee")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == nm.TIPO_EVENTO and evento.get("ORIGEN") == "SaaS (administracion)"
            and evento.get("ID_ENTIDAD") == "C3" and "->" in str(evento.get("DETALLE")),
            f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    revisar(b"Guardado en el Excel en" in admin.get("/nodo-integral").data, "la pantalla muestra el aviso con el tiempo")
    huella_antes = huella(COPIA_TEMPORAL)
    dato(admin, ficha_admin, "tarifa_km", str(nueva_tarifa))
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "el mismo valor otra vez: SIN_CAMBIO, no se escribe nada")
    nueva_accion = "Mantener" if str(manuales["meta_ads"]["valor"]).strip() == "Subir" else "Subir"
    dato(admin, ficha_admin, "meta_ads", nueva_accion)
    fila_ads = next(f for f in nm.leer_nodo() if f["campo"] == "meta_ads")
    revisar(fila_ads["valor"] == nueva_accion and fila_ads["definido"],
            f"Accion Meta Ads {manuales['meta_ads']['valor']!r} -> {nueva_accion!r} (celda C8)")
    dato(admin, ficha_admin, "tarifa_km", str(tarifa_inicial))
    despues = {f["celda"]: f["valor"] for f in nm.leer_nodo()}
    revisar(all(despues[f["celda"]] == f["valor"] for f in filas_nodo if f["celda"] != "C8"),
            f"la tarifa volvio a {tarifa_inicial}; las demas filas de la hoja nunca cambiaron")

    # "Oferta hoy": C16 'Activacion de Promocion' la escribe el robot (#84)
    oferta = ap.leer_activacion()
    for nombre_cliente, cliente in (("despacho", despacho), ("administracion", admin)):
        pagina = cliente.get("/nodo-integral")
        revisar(b"Oferta hoy" in pagina.data and b"Actualizar indicador de oferta" not in pagina.data,
                f"{nombre_cliente} ve el indicador 'Oferta hoy' y no hay boton para actualizarlo")
    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = admin.post("/nodo-integral/oferta", data={"ficha": ficha_admin})
    revisar(respuesta.status_code == 404 and huella(COPIA_TEMPORAL) == huella_antes,
            "el SaaS no tiene forma de escribir C16: esa direccion no existe (404) y el Excel no cambia")
    ap.actualizar_activacion_promocion()          # como lo hace la tarea diaria: el robot, sin el SaaS
    datos.olvidar_cache()
    despues_oferta = ap.leer_activacion()
    revisar(despues_oferta["valor_celda"] == oferta["valor"] and despues_oferta["al_dia"],
            f"C16 'Activacion de Promocion': {oferta['valor_celda']!r} -> {despues_oferta['valor_celda']!r} ({despues_oferta['motivo']})")
    if not oferta["al_dia"]:
        evento = ultimo_evento()
        revisar(evento.get("TIPO_EVENTO") == ap.TIPO_EVENTO and evento.get("ORIGEN") == "Robot Python"
                and evento.get("ID_ENTIDAD") == "C16", f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    huella_antes = huella(COPIA_TEMPORAL)
    resultado = ap.actualizar_activacion_promocion()
    revisar(resultado["estado"] == "SIN_CAMBIO" and huella(COPIA_TEMPORAL) == huella_antes,
            "el robot otra vez: SIN_CAMBIO, no se escribe nada")
    texto = html.unescape(despacho.get("/nodo-integral").get_data(as_text=True))
    revisar(f'sello--grande">{despues_oferta["valor_celda"]}<' in texto
            and all(c["codigo"] in texto for c in despues_oferta["combos"]),
            f"despues de la corrida del robot la pantalla dice Oferta hoy: {despues_oferta['valor_celda']}, con el detalle")

    # ------------------------------------------------------------------ M
    paso("M) Tablas cruce maestras: solo lectura para los dos usuarios")
    for ruta, hoja, nombre_tabla, llave in (("/tabla-cruce-maestra", "TABLA CRUCE MAESTRA", "TablaProductos", "Nombre tecnico completo"),
                                            ("/tabla-cruce-maestra-mp", "TABLA CRUCE MAESTRA MP", "TablaCruceMaestraMP", "Materia Prima")):
        ws, tabla = obtener_tabla(leer_excel(), hoja, nombre_tabla)
        registros = [r for r in leer_tabla(ws, tabla) if str(r.get(llave) or "").strip()]
        nombres = [str(r[llave]).strip() for r in registros]
        codigos = [str(r.get("Codigo") or "").strip() for r in registros]
        for nombre_cliente, cliente in (("despacho", despacho), ("administracion", admin)):
            pagina = cliente.get(ruta)
            texto = html.unescape(pagina.get_data(as_text=True))
            revisar(pagina.status_code == 200 and all(n in texto for n in nombres) and all(c in texto for c in codigos),
                    f"{nombre_cliente} ve los {len(nombres)} productos de '{hoja}' con su codigo")
            revisar(b"data-escribe" not in pagina.data, f"'{hoja}' ({nombre_cliente}): sin formularios, es solo lectura")

    # ------------------------------------------------------------------ N
    paso("N) Tabla combos dinamicos: lectura para todos; configuracion solo administracion (#65)")

    def leer_combos():
        return {c["codigo"]: c for c in cc.leer_combos()}

    def guardar_combo(cliente, ficha, combo, **cambios):
        formulario = {"codigo": combo["codigo"], "tipo": combo["tipo"], "operador": combo["operador"],
                      "temperatura": f"{combo['valor']:g}" if isinstance(combo["valor"], (int, float)) else "",
                      "lluvia": combo["valor"] if combo["valor"] in cc.CLIMAS else "",
                      "prioridad": str(combo["prioridad"]), "mensaje": combo["mensaje"],
                      "descuento": f"{combo['descuento']:g}", "ficha": ficha}
        formulario.update(cambios)
        return cliente.post("/combos/guardar", data=formulario)

    combos_inicio = leer_combos()
    primero = next(iter(combos_inicio.values()))
    pagina = despacho.get("/combos")
    texto = html.unescape(pagina.get_data(as_text=True))
    revisar(pagina.status_code == 200 and all(c["nombre"] in texto and c["codigo"] in texto and c["mensaje"] in texto
                                              for c in combos_inicio.values()),
            f"despacho ve los {len(combos_inicio)} combos con su codigo, estado, precio y mensaje")
    revisar(b"data-escribe" not in pagina.data and b'name="codigo"' not in pagina.data,
            "despacho no ve los formularios de configuracion")
    revisar("el clima no se actualiza" in texto and str(leer_excel()["NODO INTEGRAL"]["C5"].value).strip() in texto,
            "la pantalla muestra el clima de la hoja y avisa que en la copia de demostracion no se actualiza")
    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = guardar_combo(despacho, ficha_despacho, primero, descuento="0")
    revisar(respuesta.status_code == 403 and huella(COPIA_TEMPORAL) == huella_antes,
            "despacho no puede cambiarlos: el servidor responde 403 y el Excel no cambia")
    pagina = admin.get("/combos")
    revisar(pagina.data.count(b'name="codigo"') == len(combos_inicio)
            and all(f"Guardar {c['nombre']}".encode() in pagina.data for c in combos_inicio.values()),
            f"administracion ve un formulario por combo ({len(combos_inicio)})")
    for cambio in ({"tipo": "FERIADO"}, {"tipo": "DIA_PAGO"}, {"descuento": "30"}, {"mensaje": "=1+1"},
                   {"tipo": "TEMPERATURA", "temperatura": "80"}, {"prioridad": "0"}, {"operador": "<>"},
                   {"tipo": "LLUVIA", "lluvia": "Nublado"}, {"codigo": "Baguette"}):
        guardar_combo(admin, ficha_admin, primero, **cambio)
    pagina = admin.get("/combos")
    revisar("El robot no lo permitió".encode() in pagina.data and huella(COPIA_TEMPORAL) == huella_antes,
            "detonador FERIADO y DIA_PAGO, descuento 30, mensaje con '=', temperatura 80, prioridad 0, operador '<>', "
            "clima 'Nublado' y un combo que no existe: rechazados, no se escribe nada")
    descuento_nuevo = 5 if primero["descuento"] != 5 else 7.5
    guardar_combo(admin, ficha_admin, primero, descuento=f"{descuento_nuevo:g}")
    pagina = admin.get("/combos")
    ahora = leer_combos()[primero["codigo"]]
    revisar(ahora["descuento"] == descuento_nuevo and round(ahora["precio"], 2) != round(primero["precio"], 2),
            f"{primero['nombre']}: descuento {primero['descuento']:g} % -> {descuento_nuevo:g} % (celda {primero['celdas']['descuento']}); "
            f"el Excel recalculo el precio {primero['precio']:,.2f} -> {ahora['precio']:,.2f}")
    revisar(b"Guardado en el Excel en" in pagina.data, "la pantalla muestra el aviso con el tiempo")
    ws, tabla = obtener_tabla(leer_excel(), "RegistroEventos", "tblRegistroEventos")
    evento = [e for e in leer_tabla(ws, tabla) if e.get("TIPO_EVENTO") == cc.TIPO_EVENTO][-1]
    revisar(evento.get("ORIGEN") == "SaaS (administracion)" and evento.get("ID_ENTIDAD") == primero["codigo"]
            and "->" in str(evento.get("DETALLE")), f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    huella_antes = huella(COPIA_TEMPORAL)
    guardar_combo(admin, ficha_admin, ahora)
    revisar(huella(COPIA_TEMPORAL) == huella_antes and "ya era esa".encode() in admin.get("/combos").data,
            "la misma configuracion otra vez: SIN_CAMBIO, no se escribe nada")
    clima_hoy = str(leer_excel()["NODO INTEGRAL"]["C5"].value or "").strip()
    inactivo = next((c for c in leer_combos().values() if not c["activo"]), None)
    if inactivo is None or clima_hoy not in cc.CLIMAS:
        revisar(clima_hoy in cc.CLIMAS, f"'NODO INTEGRAL' C5 dice {clima_hoy!r} (se esperaba Lluvia o Despejado)")
        print("       (no hay un combo INACTIVO en esta copia: no se prueba el detonador LLUVIA en la pantalla)")
    else:
        guardar_combo(admin, ficha_admin, inactivo, tipo="LLUVIA", lluvia=clima_hoy)
        cambiado = leer_combos()[inactivo["codigo"]]
        revisar(cambiado["activo"] and cambiado["tipo"] == "LLUVIA" and cambiado["valor"] == clima_hoy
                and cambiado["clima"] == clima_hoy,
                f"{inactivo['nombre']} con detonador LLUVIA = {clima_hoy} (el clima de hoy en la hoja): pasa a ACTIVO "
                f"y el robot escribe {clima_hoy!r} en la columna B")
        texto = html.unescape(despacho.get("/combos?actualizar=1").get_data(as_text=True))
        revisar(f"Clima = {clima_hoy}" in texto, "despacho ve el detonador nuevo en la tabla")
    for codigo, original in combos_inicio.items():
        if any(leer_combos()[codigo][c] != original[c] for c in cc.CAMPOS):
            guardar_combo(admin, ficha_admin, original)
    combos_final = leer_combos()
    revisar(all(combos_final[k][c] == combos_inicio[k][c] for k in combos_inicio for c in cc.CAMPOS)
            and all(combos_final[k]["activo"] == combos_inicio[k]["activo"]
                    and round(combos_final[k]["precio"], 2) == round(combos_inicio[k]["precio"], 2) for k in combos_inicio),
            "se devolvio la configuracion: las celdas G a L, el estado y el precio de cada combo son los del inicio")
    revisar(ap.leer_activacion()["al_dia"], "C16 'Activacion de Promocion' quedo al dia despues de los cambios")

    # ------------------------------------------------------------------ O
    paso("O) Composicion de los combos: lectura para todos; cambios solo administracion (#85)")

    def composicion_de(codigo):
        return next(c for c in co.leer_composicion()["combos"] if c["codigo"] == codigo)

    def guardar_composicion(cliente, ficha, leido_combo, **cambios):
        propio, reventa = leido_combo["componentes"]
        formulario = {"combo": leido_combo["codigo"], "producto_1": propio["codigo_producto"],
                      "cantidad_1": f"{propio['cantidad']:g}", "producto_2": reventa["codigo_producto"],
                      "cantidad_2": f"{reventa['cantidad']:g}", "ficha": ficha}
        formulario.update(cambios)
        return cliente.post("/combos/composicion/guardar", data=formulario)

    leido = co.leer_composicion()
    combo_inicio = next(c for c in leido["combos"] if c["ajustable"])
    codigo_combo = combo_inicio["codigo"]
    propio_inicio, reventa_inicio = combo_inicio["componentes"]
    otro = next(p for p in leido["terceros"] if p["nombre"] != reventa_inicio["producto"])
    pagina = despacho.get("/combos?actualizar=1")
    texto = html.unescape(pagina.get_data(as_text=True))
    revisar("Qué lleva cada combo" in texto and all(parte["producto"] in texto for c in leido["combos"] for parte in c["componentes"]),
            "despacho ve que lleva cada combo (producto, cantidad, combos vendidos y unidades rebajadas)")
    revisar(b"/combos/composicion/guardar" not in pagina.data and b'name="producto_1"' not in pagina.data,
            "despacho no ve los formularios de composicion")
    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = guardar_composicion(despacho, ficha_despacho, combo_inicio, cantidad_1="9")
    revisar(respuesta.status_code == 403 and huella(COPIA_TEMPORAL) == huella_antes,
            "despacho no puede cambiarla: el servidor responde 403 y el Excel no cambia")
    pagina = admin.get("/combos")
    ajustables = [c for c in leido["combos"] if c["ajustable"]]
    revisar(pagina.data.count(b'name="producto_1"') == len(ajustables)
            and all(f"Guardar composición de {c['nombre']}".encode() in pagina.data for c in ajustables)
            and pagina.data.count(b'name="codigo"') == len(leido["combos"]),
            f"administracion ve un formulario de composicion por combo ({len(ajustables)}), ademas de los de configuracion")
    for cambio in ({"producto_1": reventa_inicio["codigo_producto"]}, {"producto_2": propio_inicio["codigo_producto"]},
                   {"producto_2": "Pan de yuca"}, {"cantidad_1": "0"}, {"cantidad_2": "11"}, {"cantidad_1": "2.5"},
                   {"combo": "co99"}):
        guardar_composicion(admin, ficha_admin, combo_inicio, **cambio)
    pagina = admin.get("/combos")
    revisar("El robot no lo permitió".encode() in pagina.data and huella(COPIA_TEMPORAL) == huella_antes,
            "componente 1 de reventa, componente 2 propio, producto que no existe, cantidades 0, 11 y 2.5 y un combo "
            "que no existe: rechazados, no se escribe nada")
    salidas_antes = tuple(inventario(producto, "Salidas en Combos")
                          for producto in (propio_inicio["producto"], reventa_inicio["producto"], otro["nombre"]))
    cantidad_nueva = propio_inicio["cantidad"] + 1 if propio_inicio["cantidad"] < co.CANTIDAD_MAXIMA else propio_inicio["cantidad"] - 1
    guardar_composicion(admin, ficha_admin, combo_inicio, cantidad_1=f"{cantidad_nueva:g}")
    pagina = admin.get("/combos")
    ahora = composicion_de(codigo_combo)
    propio, reventa = ahora["componentes"]
    revisar(propio["cantidad"] == cantidad_nueva and propio["congelados"] == propio_inicio["vendidos"]
            and propio["rebajadas"] == propio_inicio["unidades"] and propio["unidades"] == propio_inicio["unidades"]
            and round(ahora["precio"], 2) != round(combo_inicio["precio"], 2)
            and ahora["detalle"] == co.texto_detalle(cantidad_nueva, propio["producto"], reventa["cantidad"], reventa["producto"]),
            f"{ahora['nombre']}: {propio_inicio['cantidad']:g} -> {cantidad_nueva:g} {propio['producto']} (celda {propio['celdas']['cantidad']}); "
            f"congelados {propio['congelados']:g} combos y {propio['rebajadas']:g} unidades; precio {combo_inicio['precio']:,.2f} -> "
            f"{ahora['precio']:,.2f}; detalle '{ahora['detalle']}'")
    revisar(b"Guardado en el Excel en" in pagina.data, "la pantalla muestra el aviso con el tiempo")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == co.TIPO_EVENTO and evento.get("ORIGEN") == "SaaS (administracion)"
            and evento.get("ID_ENTIDAD") == codigo_combo and "->" in str(evento.get("DETALLE")),
            f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    huella_antes = huella(COPIA_TEMPORAL)
    guardar_composicion(admin, ficha_admin, ahora)
    revisar(huella(COPIA_TEMPORAL) == huella_antes and "ya era esa".encode() in admin.get("/combos").data,
            "la misma composicion otra vez: SIN_CAMBIO, no se escribe nada")
    guardar_composicion(admin, ficha_admin, ahora, producto_2=otro["codigo"], cantidad_2="2")
    cambiado = composicion_de(codigo_combo)
    reventa_nueva = cambiado["componentes"][1]
    historia = [h for h in cambiado["historia"] if h["producto"] == reventa_inicio["producto"]]
    revisar(reventa_nueva["producto"] == otro["nombre"] and reventa_nueva["cantidad"] == 2 and reventa_nueva["unidades"] == 0
            and (not reventa_inicio["unidades"] or (len(historia) == 1 and historia[0]["unidades"] == reventa_inicio["unidades"])),
            f"reventa {reventa_inicio['producto']} -> 2 x {otro['nombre']}: la fila vigente empieza en 0"
            + (f" y lo ya rebajado ({reventa_inicio['unidades']:g}) queda en la fila de historia {historia[0]['fila']}"
               if reventa_inicio["unidades"] and historia else " (no habia nada que congelar)"))
    texto = html.unescape(despacho.get("/combos?actualizar=1").get_data(as_text=True))
    revisar(cambiado["detalle"] in texto and ("Historia (ya no lo lleva)" in texto) == bool(cambiado["historia"]),
            "despacho ve la composicion nueva en la pantalla")
    guardar_composicion(admin, ficha_admin, combo_inicio)
    final = composicion_de(codigo_combo)
    revisar([(c["producto"], c["cantidad"]) for c in final["componentes"]]
            == [(c["producto"], c["cantidad"]) for c in combo_inicio["componentes"]]
            and round(final["precio"], 2) == round(combo_inicio["precio"], 2)
            and tuple(inventario(producto, "Salidas en Combos")
                      for producto in (propio_inicio["producto"], reventa_inicio["producto"], otro["nombre"])) == salidas_antes,
            f"se devolvio la composicion ({final['detalle']}): el precio y las salidas en combos del inventario son los del inicio")
    revisar([c["codigo"] for c in cc.leer_combos()] == list(combos_inicio),
            "los combos siguen siendo los mismos: las filas de historia no se confunden con combos")

    # ------------------------------------------------------------------ P
    paso("P) Hojas de costo: lectura para todos; precio de referencia solo administracion (Regla #9)")

    def costeo_de(codigo):
        return next(c for c in cs.leer_costeo(leer_excel()) if c["codigo"] == codigo)

    def precio_tcm(codigo):
        ws, tabla = obtener_tabla(leer_excel(), "TABLA CRUCE MAESTRA", "TablaProductos")
        return round(float(next(f for f in leer_tabla(ws, tabla) if str(f.get("Codigo")).strip() == codigo)["Precio"]), 2)

    def combo_con(producto):
        return next(c for c in co.leer_composicion()["combos"] if any(parte["producto"] == producto for parte in c["componentes"]))

    def guardar_precio(cliente, ficha, codigo, precio):
        return cliente.post("/costo/precio-referencia", data={"producto": codigo, "precio": precio, "ficha": ficha})

    pantallas_costo = (("/costo-pan-baguette", "P01"), ("/costo-cupcakes", "Pa01"))
    for ruta, codigo in pantallas_costo:
        hoja_costo = costeo_de(codigo)
        elementos = [f["elemento"] for s in hoja_costo["secciones"] for f in s["filas"]]
        pagina = despacho.get(f"{ruta}?actualizar=1")
        texto = html.unescape(pagina.get_data(as_text=True))
        revisar(pagina.status_code == 200 and all(e in texto for e in elementos)
                and f"₡{hoja_costo['costo_unitario']:,.2f}" in texto and f"₡{hoja_costo['precio_referencia']:,.2f}" in texto,
                f"despacho ve '{hoja_costo['hoja']}': {len(elementos)} elementos del costo, costo unitario "
                f"{hoja_costo['costo_unitario']:,.2f} y precio de referencia {hoja_costo['precio_referencia']:,.2f}")
        revisar(b"data-escribe" not in pagina.data and b'name="precio"' not in pagina.data,
                f"'{hoja_costo['hoja']}' (despacho): sin formulario, es solo lectura")
        pagina = admin.get(ruta)
        revisar(pagina.data.count(b'name="precio"') == 1 and "Guardar precio de referencia".encode() in pagina.data
                and f"celda {hoja_costo['celda_precio']}".encode() in pagina.data,
                f"administracion ve el formulario del precio de referencia (celda {hoja_costo['celda_precio']})")
    revisar(b"4 Costeo" in despacho.get("/costo-pan-baguette").data, "el menu muestra el grupo '4 Costeo'")

    # El cambio se prueba con un producto cuyo costo unitario este por debajo de su precio de referencia.
    # (La seccion K deja cargada una hornada de 10 Baguette: con tan pocas unidades su costo unitario
    # queda por encima del precio y el robot, con razon, no deja escribir un precio menor que ese costo.)
    hojas_costo = cs.leer_costeo(leer_excel())
    con_margen = [c for c in hojas_costo if c["costo_unitario"] is not None and c["precio_referencia"] is not None
                  and c["costo_unitario"] + 10 < c["precio_referencia"]]
    revisar(bool(con_margen), "hay al menos un producto con el costo unitario por debajo de su precio de referencia")
    if not con_margen:
        return 1
    for sin_margen in (c for c in hojas_costo if c not in con_margen and c["costo_unitario"] > c["precio_referencia"] + 10):
        huella_antes = huella(COPIA_TEMPORAL)
        guardar_precio(admin, ficha_admin, sin_margen["codigo"], f"{sin_margen['precio_referencia'] + 10:.2f}")
        pagina = admin.get(dict((c, r) for r, c in pantallas_costo)[sin_margen["codigo"]])
        revisar(b"BLOQUEADO" in pagina.data and huella(COPIA_TEMPORAL) == huella_antes,
                f"{sin_margen['nombre']}: la hornada cargada es pequeña y su costo unitario ({sin_margen['costo_unitario']:,.2f}) "
                f"supera al precio de referencia ({sin_margen['precio_referencia']:,.2f}): el robot bloquea un precio que siga "
                "bajo el costo y no escribe nada")
    inicio = con_margen[0]
    codigo_p = inicio["codigo"]
    ruta_p = dict((c, r) for r, c in pantallas_costo)[codigo_p]
    precio_inicio, costo_inicio = inicio["precio_referencia"], inicio["costo_unitario"]
    combo_inicio_p = combo_con(inicio["nombre"])
    huella_antes = huella(COPIA_TEMPORAL)
    respuesta = guardar_precio(despacho, ficha_despacho, codigo_p, f"{precio_inicio + 10:.2f}")
    revisar(respuesta.status_code == 403 and huella(COPIA_TEMPORAL) == huella_antes,
            "despacho no puede cambiar el precio de referencia: el servidor responde 403 y el Excel no cambia")
    bajo_el_costo = f"{max(costo_inicio - 1, 1):.2f}"
    for codigo, precio in ((codigo_p, bajo_el_costo), (codigo_p, "0"), (codigo_p, "-5"), (codigo_p, "abc"), ("L01", "900"),
                           ("Zz99", "900")):
        guardar_precio(admin, ficha_admin, codigo, precio)
    pagina = admin.get(ruta_p)
    revisar("El robot no lo permitió".encode() in pagina.data and b"BLOQUEADO" in pagina.data
            and huella(COPIA_TEMPORAL) == huella_antes,
            f"{inicio['nombre']}: precio bajo el costo unitario ({bajo_el_costo}), 0, negativo, texto, un producto de reventa "
            "(L01) y un codigo que no existe: rechazados, no se escribe nada")
    guardar_precio(admin, ficha_admin, codigo_p, "")
    revisar("Escribe el precio de referencia".encode() in admin.get(ruta_p).data
            and huella(COPIA_TEMPORAL) == huella_antes, "precio vacio: no se escribe nada")

    precio_nuevo = round(precio_inicio + 10, 2)
    guardar_precio(admin, ficha_admin, codigo_p, f"{precio_nuevo:.2f}")
    pagina = admin.get(ruta_p)
    ahora = costeo_de(codigo_p)
    combo_ahora = combo_con(inicio["nombre"])
    revisar(ahora["precio_referencia"] == precio_nuevo and round(ahora["costo_unitario"], 2) == round(costo_inicio, 2)
            and round(ahora["utilidad"], 2) == round(precio_nuevo - costo_inicio, 2),
            f"{inicio['nombre']}: precio de referencia {precio_inicio:,.2f} -> {precio_nuevo:,.2f} (celda {inicio['celda_precio']}); "
            f"el costo unitario no cambia ({costo_inicio:,.2f}) y la utilidad pasa a {ahora['utilidad']:,.2f}")
    revisar(precio_tcm(codigo_p) == precio_nuevo and round(combo_ahora["precio"], 2) > round(combo_inicio_p["precio"], 2),
            f"el Excel recalculo: 'TABLA CRUCE MAESTRA' lee {precio_tcm(codigo_p):,.2f} y {combo_ahora['nombre']} sube de "
            f"{combo_inicio_p['precio']:,.2f} a {combo_ahora['precio']:,.2f}")
    revisar(b"Guardado en el Excel en" in pagina.data, "la pantalla muestra el aviso con el tiempo")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == "CONFIG_PRECIO_REFERENCIA" and evento.get("ORIGEN") == "SaaS (administracion)"
            and evento.get("ID_ENTIDAD") == codigo_p and "->" in str(evento.get("DETALLE")),
            f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    huella_antes = huella(COPIA_TEMPORAL)
    guardar_precio(admin, ficha_admin, codigo_p, f"{precio_nuevo:.2f}")
    revisar(huella(COPIA_TEMPORAL) == huella_antes and "ya era".encode() in admin.get(ruta_p).data,
            "el mismo precio otra vez: SIN_CAMBIO, no se escribe nada")
    texto = html.unescape(despacho.get(f"{ruta_p}?actualizar=1").get_data(as_text=True))
    revisar(f"₡{precio_nuevo:,.2f}" in texto, "despacho ve el precio de referencia nuevo en la pantalla")
    guardar_precio(admin, ficha_admin, codigo_p, f"{precio_inicio:.2f}")
    final = costeo_de(codigo_p)
    revisar(final["precio_referencia"] == precio_inicio and precio_tcm(codigo_p) == round(precio_inicio, 2)
            and round(combo_con(inicio["nombre"])["precio"], 2) == round(combo_inicio_p["precio"], 2),
            f"se devolvio el precio de referencia ({precio_inicio:,.2f}): la tabla cruce maestra y el combo vuelven a su valor")
    texto = html.unescape(admin.get("/nodo-integral").get_data(as_text=True))
    revisar("Es una prevista" in texto and "todavía no hay ninguna campaña" in texto,
            "NODO INTEGRAL: la nota de C8 'Accion Meta Ads' dice que es una prevista (#96)")

    # ------------------------------------------------------------------ F
    paso("F) Cierre")
    revisar(not buscar_errores(COPIA_TEMPORAL.read_bytes()), "la copia temporal quedo con 0 errores de formula")
    revisar(huella(COPIA_LOCAL) == huella_local_antes, "datos_prueba\\maestro_prueba.xlsx quedo IGUAL (no se toco)")

    print()
    if errores:
        print(f"RESULTADO: {len(errores)} linea(s) ERROR")
        return 1
    print("RESULTADO: TODO OK (0 lineas ERROR)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
