"""Prueba de Pedidos digitales SOBRE LA COPIA LOCAL del Excel (no toca Google Drive).

OJO: Google Maps SI se consulta de verdad (3 llamadas; entran en la cuota
gratis). Necesita internet y GOOGLE_MAPS_API_KEY en el .env.

Usa la misma copia que la prueba de Ventas despacho:
    datos_prueba/maestro_prueba.xlsx

Version 27-sep-2026 (#7, Regla #30): forma de pago OBLIGATORIA (Efectivo /
SINPE Movil) y anulacion de la COMANDA COMPLETA (anular_pedido_digital).
Escribe 3 pedidos (Canas en Efectivo, retiro en Efectivo, retiro por SINPE)
y 1 anulacion (la del pedido de Canas).

Version 29-sep-2026 (m013, Sesion A del bot): el pedido de Canas lleva el
TELEFONO del cliente (#33 opcion B); se prueba marcar_entregado (recibo al
entregar, opcion C), que un pedido entregado ya no se anula y que el numero
SINPE se lee de NODO INTEGRAL. Necesita la copia local CON la m013 aplicada.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo1_captura_operativa.probar_pedidos_digitales
"""
import os

# Forzar modo local ANTES de importar el resto (asi nunca toca Drive)
os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import logging                                                   # noqa: E402
import time as reloj                                             # noqa: E402

from conexion_drive import leer_excel                            # noqa: E402
from conexion_drive.tablas import columnas_de_tabla, leer_tabla  # noqa: E402
from .catalogo import Catalogo, ComboInactivo, ProductoNoEncontrado, leer_reglas_del_dia  # noqa: E402
from .pedidos_digitales import (  # noqa: E402
    AnulacionNoPermitida, EntregaNoPermitida, FueraDeZona, MetodoPagoInvalido, anular_linea_pedido_digital,
    anular_pedido_digital, cotizar_envio, cotizar_pedido, marcar_entregado, registrar_pedido_digital,
)

CANAS = (10.4318, -85.0983)       # cliente de prueba dentro de Canas
LIBERIA = (10.6346, -85.4377)     # cliente de prueba fuera de zona
TELEFONO_TWILIO = "whatsapp:+50688887777"   # asi lo manda Twilio (numero supuesto)
TELEFONO_LIMPIO = "+50688887777"            # asi debe quedar en la columna TELEFONO

logging.basicConfig(level=logging.WARNING)


def paso(titulo):
    print(f"\n--- {titulo} ---")


def suma_columna_flujo(columna):
    """Suma de una columna de 'Flujo caja diario' en todo el ano (3 = C, 11 = K)."""
    ws = leer_excel()["Flujo caja diario"]
    return round(sum(ws.cell(f, columna).value or 0 for f in range(5, ws.max_row + 1)
                     if isinstance(ws.cell(f, columna).value, (int, float))), 2)


def ventas_digitales_hoy():
    """C = Ventas digitales (solo comanda)."""
    return suma_columna_flujo(3)


def envios_cobrados():
    """K = Envios cobrados (columna nueva de la m011)."""
    return suma_columna_flujo(11)


def celda_pedido(id_pedido, columna):
    """Valor de una columna de 'Interface digital' para un pedido (lee la copia local)."""
    wb = leer_excel()
    ws = wb["Interface digital"]
    tabla = ws.tables["TablaInterfaceDigital"]
    letra = columnas_de_tabla(tabla)[columna]
    for fila in leer_tabla(ws, tabla):
        if fila.get("ID_PEDIDO") == id_pedido:
            return ws[f"{letra}{fila['_fila']}"].value
    return "NO EXISTE"


def debe_rechazar(descripcion, funcion, *args, error=AnulacionNoPermitida, **kwargs):
    """Corre la funcion y espera un rechazo. Si NO la rechaza, avisa ERROR."""
    try:
        funcion(*args, **kwargs)
        print(f"  ERROR: {descripcion} NO fue rechazado")
    except error as rechazo:
        print(f"  Rechazado correctamente: {rechazo}")


if __name__ == "__main__":
    id_prueba = f"PRUEBA-{int(reloj.time())}"      # simula el ID unico del mensaje de WhatsApp
    antes = ventas_digitales_hoy()
    envios_antes = envios_cobrados()
    print(f"Ventas digitales en Flujo caja ANTES: {antes}   Envios cobrados ANTES: {envios_antes}")

    catalogo = Catalogo(leer_excel())
    combo = next((c for c in catalogo.combos if c["activo"]), None)

    paso("1a) El cliente comparte su ubicacion (Canas) -> cotizar_envio")
    cot = cotizar_envio(CANAS)
    print(f"  {cot}")

    paso("1b) Antes de confirmar: TOTAL COMPLETO que ve el cliente -> cotizar_pedido")
    lineas = [{"producto_id": "P01", "cantidad": 3}, {"producto_id": "Pa01", "cantidad": 2}]
    if combo:
        lineas.append({"producto_id": combo["codigo"], "cantidad": 1})
    resumen = cotizar_pedido(lineas, distancia_km=cot["km"], minutos_entrega=cot["minutos"])
    for cantidad, nombre, subtotal in resumen["lineas"]:
        print(f"  {cantidad} x {nombre}: {subtotal}")
    print(f"  Envio ({resumen['km']} km, {resumen['minutos']} min): {resumen['envio']}")
    print(f"  TOTAL A PAGAR: {resumen['total']}")

    paso("1c) El cliente dice 'Si' -> registrar_pedido_digital (NO vuelve a llamar a Google)")
    resultado = registrar_pedido_digital(
        lineas, id_mensaje=id_prueba, cliente="Cliente de prueba", ubicacion=CANAS,
        distancia_km=cot["km"], minutos_entrega=cot["minutos"], total_cotizado=resumen["total"],
        metodo_pago="Efectivo", telefono=TELEFONO_TWILIO,
    )
    for clave, valor in resultado.items():
        print(f"  {clave}: {valor}")
    print("  " + ("OK: total registrado = total mostrado" if resultado["total"] == resumen["total"]
                  else "ERROR: los totales no coinciden"))
    print("  " + ("OK: la comanda trae la forma de pago (Efectivo)" if resultado.get("metodo_pago") == "Efectivo"
                  else "ERROR: la comanda no trae la forma de pago"))
    tel_excel = celda_pedido(resultado["id_pedido"], "TELEFONO")
    print("  " + (f"OK: TELEFONO guardado limpio en el Excel ({tel_excel})" if tel_excel == TELEFONO_LIMPIO
                  else f"ERROR: TELEFONO en el Excel = {tel_excel!r} (se esperaba {TELEFONO_LIMPIO})"))

    paso("2) El mismo mensaje otra vez (debe decir DUPLICADO)")
    print(" ", registrar_pedido_digital(lineas, id_mensaje=id_prueba, metodo_pago="Efectivo"))

    paso("3) Producto inexistente (debe rechazarse sin escribir nada)")
    try:
        registrar_pedido_digital([{"producto": "PAN INVENTADO", "cantidad": 1}], id_mensaje=id_prueba + "-b",
                                 metodo_pago="Efectivo")
    except ProductoNoEncontrado as error:
        print(f"  Rechazado correctamente: {error}")

    paso("4) Combo INACTIVO (debe rechazarse sin escribir nada)")
    inactivo = next((c for c in catalogo.combos if not c["activo"]), None)
    if inactivo:
        try:
            registrar_pedido_digital([{"producto": inactivo["codigo"], "cantidad": 1}], id_mensaje=id_prueba + "-c",
                                     metodo_pago="Efectivo")
        except ComboInactivo as error:
            print(f"  Rechazado correctamente: {error}")

    paso("4b) Cliente en Liberia (fuera de zona): el bot avisa y no se registra")
    cot_lejos = cotizar_envio(LIBERIA)
    print(f"  cotizar_envio: dentro_zona={cot_lejos['dentro_zona']}, km={cot_lejos['km']}")
    try:
        registrar_pedido_digital(lineas, id_mensaje=id_prueba + "-d", ubicacion=LIBERIA, metodo_pago="Efectivo")
    except FueraDeZona as error:
        print(f"  Rechazado correctamente: {error}")

    paso("4c) Retiro en el local (sin ubicacion): envio 0")
    retiro = registrar_pedido_digital([{"producto": "P01", "cantidad": 1}], id_mensaje=id_prueba + "-e",
                                      metodo_pago="Efectivo")
    print(f"  {retiro['id_pedido']}: envio {retiro['envio']}, total {retiro['total']}")

    paso("4d) Pedido SIN forma de pago (debe rechazarse sin escribir nada)")
    debe_rechazar("el pedido sin forma de pago", registrar_pedido_digital,
                  [{"producto": "P01", "cantidad": 1}], id_mensaje=id_prueba + "-g", error=MetodoPagoInvalido)

    paso("5a) Intentar anular UNA linea (ya no se permite: debe rechazarse)")
    debe_rechazar("anular una linea", anular_linea_pedido_digital, resultado["id_pedido"], "Pa01")

    paso("5b) La administracion anula la COMANDA COMPLETA del pedido de Canas (Efectivo)")
    anulacion = anular_pedido_digital(resultado["id_pedido"], "Prueba: el cliente no recibio el pedido")
    for clave, valor in anulacion.items():
        print(f"  {clave}: {valor}")
    print("  " + ("OK: se revirtio el total completo (comanda + envio)"
                  if anulacion["total_revertido"] == -resultado["total"] else "ERROR: el monto revertido no cuadra"))

    paso("5c) Anular OTRA VEZ el mismo pedido (debe decir que ya estaba anulado)")
    debe_rechazar("la segunda anulacion", anular_pedido_digital, resultado["id_pedido"], "Prueba repetida")

    paso("5d) Anular la fila de anulacion (debe rechazarse)")
    debe_rechazar("anular la anulacion", anular_pedido_digital, anulacion["id_anulacion"], "Prueba")

    paso("5e) Anular sin motivo (debe rechazarse)")
    debe_rechazar("anular sin motivo", anular_pedido_digital, retiro["id_pedido"], "   ")

    paso("5f) Pedido de retiro pagado por SINPE Movil, luego intentar anularlo (pago en firme)")
    sinpe = registrar_pedido_digital([{"producto": "P01", "cantidad": 1}], id_mensaje=id_prueba + "-f",
                                     metodo_pago="SINPE M\u00f3vil")
    print(f"  {sinpe['id_pedido']}: registrado, pago {sinpe['metodo_pago']}, total {sinpe['total']}")
    debe_rechazar("anular un pedido SINPE", anular_pedido_digital, sinpe["id_pedido"], "Prueba SINPE")

    paso("6a) Numero SINPE de la panaderia (NODO INTEGRAL, m013)")
    sinpe_nodo = leer_reglas_del_dia(leer_excel())["sinpe_movil"]
    print("  " + (f"OK: SINPE = {sinpe_nodo}" if sinpe_nodo == "8888-8888"
                  else f"ERROR: SINPE = {sinpe_nodo!r} (se esperaba '8888-8888')"))

    paso("6b) El cliente retira el pedido 4c -> marcar_entregado (datos del recibo)")
    entrega = marcar_entregado(retiro["id_pedido"], origen="Prueba")
    print(f"  estado: {entrega['estado']}, entregado_en: {entrega['entregado_en']}")
    for cantidad, nombre, subtotal in entrega["lineas"]:
        print(f"  recibo: {cantidad} x {nombre}: {subtotal}")
    print(f"  recibo: total {entrega['total']}, pago {entrega['metodo_pago']}, cliente {entrega['cliente']!r}, "
          f"telefono {entrega['telefono']!r}")
    ok = (entrega["estado"] == "ENTREGADO" and entrega["total"] == retiro["total"]
          and celda_pedido(retiro["id_pedido"], "ENTREGADO") not in (None, "")
          and celda_pedido(retiro["id_pedido"], "ESTADO") == "CONFIRMADO")
    print("  " + ("OK: ENTREGADO con fecha y hora, ESTADO sigue CONFIRMADO, total del recibo = total del pedido"
                  if ok else "ERROR: la entrega no quedo como se esperaba"))

    paso("6c) El repartidor manda el mensaje OTRA VEZ (debe decir YA_ENTREGADO, sin segundo recibo)")
    otra = marcar_entregado(retiro["id_pedido"], origen="Prueba")
    print("  " + (f"OK: {otra['estado']} ({otra['entregado_en']})" if otra["estado"] == "YA_ENTREGADO"
                  else f"ERROR: estado {otra['estado']}"))

    paso("6d) Anular el pedido YA ENTREGADO (regla nueva: debe rechazarse)")
    debe_rechazar("anular un pedido entregado", anular_pedido_digital, retiro["id_pedido"], "Prueba entregado")

    paso("6e) Marcar entregado el pedido de Canas, que fue ANULADO (debe rechazarse)")
    debe_rechazar("entregar un pedido anulado", marcar_entregado, resultado["id_pedido"], error=EntregaNoPermitida)

    paso("6f) Marcar entregado la fila de anulacion -A (debe rechazarse)")
    debe_rechazar("entregar una fila -A", marcar_entregado, anulacion["id_anulacion"], error=EntregaNoPermitida)

    paso("6g) Marcar entregado un pedido que no existe (debe rechazarse)")
    debe_rechazar("entregar un pedido inexistente", marcar_entregado, "WA-99999", error=LookupError)

    despues = ventas_digitales_hoy()
    envios_despues = envios_cobrados()
    esperado = round(retiro["total"] + sinpe["total"], 2)       # Canas se anulo; quedan los 2 retiros (envio 0)
    print(f"\nVentas digitales en Flujo caja DESPUES: {despues}  (diferencia {round(despues - antes, 2)}, "
          f"esperada {esperado})")
    print(f"Envios cobrados DESPUES: {envios_despues}  (diferencia {round(envios_despues - envios_antes, 2)}, "
          f"esperada 0)")
    cuadra = round(despues - antes, 2) == esperado and round(envios_despues - envios_antes, 2) == 0
    print("OK: el flujo de caja cuadra (la anulacion resto comanda y envio)" if cuadra
          else "ERROR: el flujo de caja no cuadra")
    print("Abre datos_prueba/maestro_prueba.xlsx y revisa el final de 'Interface digital' (fila -A en negativo),")
    print("'Detalle comanda digital' (filas -A con ESTADO ANULADO) y 'RegistroEventos' (ANULACION_PEDIDO).")
    print("En 'NODO INTEGRAL' C9/C10 deben estar los minutos y km del pedido de Canas.")
    print("En 'Interface digital' columna P (TELEFONO) del pedido de Canas y Q (ENTREGADO) del retiro en Efectivo;")
    print("en 'RegistroEventos' el evento ENTREGA_PEDIDO.")
