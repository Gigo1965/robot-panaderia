"""Prueba de NODO INTEGRAL SIN internet, sobre la copia local del Excel.

Parte 1: interpreta respuestas de ejemplo (clima, BCCR, Hacienda) y la regla de dias de pago.
Parte 2: escribe valores simulados en datos_prueba/maestro_prueba.xlsx y revisa que los
         combos se activen solos segun la temperatura, y que el dia de pago apague los combos.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.probar_nodo
"""
import os

os.environ["ORIGEN_EXCEL"] = "local"
os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

import logging                                   # noqa: E402
from datetime import date                        # noqa: E402

from . import fuentes                            # noqa: E402
from .nodo_integral import actualizar_nodo, combos_activos  # noqa: E402

logging.basicConfig(level=logging.WARNING)
fallos = []


def revisar(nombre, obtenido, esperado):
    ok = obtenido == esperado
    print(f"  {'OK   ' if ok else 'FALLO'} {nombre}: {obtenido!r}" + ("" if ok else f" (esperado {esperado!r})"))
    if not ok:
        fallos.append(nombre)


print("=== Parte 1: interpretar respuestas de ejemplo ===")
lluvia = {"current": {"time": "2026-09-23T15:00", "temperature_2m": 26.4, "precipitation": 1.2, "rain": 1.2,
                      "weather_code": 61}}
sol = {"current": {"time": "2026-09-23T12:00", "temperature_2m": 33.06, "precipitation": 0, "rain": 0,
                   "weather_code": 1}}
revisar("clima con lluvia", fuentes.interpretar_clima(lluvia)["estado"], "Lluvia")
revisar("clima despejado", fuentes.interpretar_clima(sol)["estado"], "Despejado")
revisar("temperatura redondeada", fuentes.interpretar_clima(sol)["temperatura"], 33.1)

hacienda = {"venta": {"fecha": "2026-09-23 00:00:00", "valor": 451.86},
            "compra": {"fecha": "2026-09-23 00:00:00", "valor": 447.19}}
tc = fuentes.interpretar_tipo_cambio_hacienda(hacienda)
revisar("tipo de cambio compra (referencia BCCR)", tc["compra"], 447.19)
revisar("tipo de cambio venta (referencia BCCR)", tc["venta"], 451.86)
revisar("fecha del tipo de cambio", tc["fecha"], "2026-09-23")
try:
    fuentes.interpretar_tipo_cambio_hacienda({"venta": {"valor": 0}, "compra": {"valor": 0}})
    revisar("rechaza un tipo de cambio absurdo", "no rechazo", "rechazado")
except fuentes.FuenteNoDisponible:
    revisar("rechaza un tipo de cambio absurdo", "rechazado", "rechazado")

for fecha, esperado in ((date(2026, 9, 15), True), (date(2026, 9, 30), True), (date(2026, 9, 23), False),
                        (date(2026, 2, 28), True), (date(2028, 2, 28), False), (date(2028, 2, 29), True),
                        (date(2026, 8, 31), False), (date(2026, 11, 15), True)):
    revisar(f"dia de pago {fecha} ({fecha.strftime('%a')})", fuentes.es_dia_de_pago(fecha), esperado)

print("\n=== Parte 2: escribir en la copia local y ver los combos ===")
base = {"tipo_cambio": 451.86, "feriado_nombre": "Ninguno", "feriado_hoy": "NO",
        "dia_pago_lista": "NO", "dia_pago_hoy": "NO"}

print("-- Dia caluroso: 33.1 C, despejado")
print("  ", actualizar_nodo(valores={**base, "clima": "Despejado", "temperatura": 33.1}, errores={}))
activos, reglas = combos_activos()
revisar("combos activos con 33.1 C", activos, ["co02 Combo 2", "co03 Combo 3"])

print("-- Dia fresco: 24.0 C, lluvia")
print("  ", actualizar_nodo(valores={**base, "clima": "Lluvia", "temperatura": 24.0}, errores={}))
activos, reglas = combos_activos()
revisar("combos activos con 24 C", activos, ["co01 Combo 1", "co04 Combo 4"])

print("-- Mismos valores otra vez (no debe subir nada)")
r = actualizar_nodo(valores={**base, "clima": "Lluvia", "temperatura": 24.0}, errores={})
revisar("sin cambios", r["hubo_cambios"], False)

print("-- Falla el tipo de cambio: se conserva el valor anterior y se anota el error")
r = actualizar_nodo(valores={"clima": "Lluvia", "temperatura": 24.0}, errores={"tipo_cambio": "BCCR sin respuesta"})
revisar("se anoto el error", list(r["errores"]), ["tipo_cambio"])

print("-- Dia de pago: 5% a todo y sin combos (Regla #8)")
actualizar_nodo(valores={**base, "clima": "Despejado", "temperatura": 33.1, "dia_pago_lista": "SI",
                         "dia_pago_hoy": "SI"}, errores={})
activos, reglas = combos_activos()
revisar("descuento general activo", reglas["descuento_general"], True)
revisar("detonador", reglas["detonador_general"], "DIA_PAGO")

actualizar_nodo(valores={**base, "clima": "Despejado", "temperatura": 28.0}, errores={})   # dejar dia normal
print("\nRESULTADO:", "TODO BIEN" if not fallos else f"{len(fallos)} FALLOS: {fallos}")
