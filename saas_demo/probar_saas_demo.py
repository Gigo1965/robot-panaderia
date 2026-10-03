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
  F) La copia local datos_prueba/maestro_prueba.xlsx quedo IGUAL (misma huella).

Necesita: LibreOffice (igual que las otras pruebas) y Flask.
Escribe 2 veces en la copia temporal (unos 10 s cada una en la PC).

Uso (desde la carpeta ROBOT PANADERO):
    py -m saas_demo.probar_saas_demo
"""
import hashlib
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
from microservicios.grupo1_captura_operativa import pedidos_digitales as pd   # noqa: E402

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
