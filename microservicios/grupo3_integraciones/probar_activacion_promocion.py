"""Prueba de activacion_promocion.py SOBRE UNA COPIA TEMPORAL (no toca Drive ni la copia local).

Copia datos_prueba/maestro_prueba.xlsx a una carpeta temporal y trabaja ahi; al terminar
la carpeta se borra sola. Por eso NO hace falta "foto antes -> prueba -> restaurar".
No usa internet: el dia de pago y los combos se simulan escribiendo en la copia temporal.

Que revisa (Decisiones Pendientes #84, 'NODO INTEGRAL' C16 'Activacion de Promocion'):
  1) Dia normal con combos activos: C16 = ON, evento con el valor anterior, el nuevo y el motivo.
  2) Otra vez sin que nada cambie: SIN_CAMBIO, no se escribe nada.
  3) Dia de pago: sigue ON, pero por el descuento general (ese dia no hay combos, Regla #8).
  4) Dia normal SIN ningun combo activo: C16 pasa a OFF.
  5) Vuelven los combos: C16 vuelve a ON.
  6) Solo cambio C16; C19 y C20 siguen siendo formulas; 0 errores de formula; copia local intacta.

Necesita LibreOffice (igual que las otras pruebas). Escribe 6 veces: tarda cerca de 1 minuto.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.probar_activacion_promocion
"""
import atexit
import hashlib
import os
import shutil
import sys
import tempfile
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent.parent
COPIA_LOCAL = RAIZ_PROYECTO / "datos_prueba" / "maestro_prueba.xlsx"
if not COPIA_LOCAL.exists():
    sys.exit(f"No encontre la copia local: {COPIA_LOCAL}\n"
             "La prueba necesita datos_prueba\\maestro_prueba.xlsx (la misma de las otras pruebas).")
_carpeta = Path(tempfile.mkdtemp(prefix="prueba_activacion_"))
COPIA_TEMPORAL = _carpeta / "copia_prueba.xlsx"
shutil.copy2(COPIA_LOCAL, COPIA_TEMPORAL)
atexit.register(shutil.rmtree, _carpeta, ignore_errors=True)

# Modo local sobre la copia temporal, ANTES de importar el robot (asi nunca toca Drive).
os.environ["ORIGEN_EXCEL"] = "local"
os.environ["EXCEL_LOCAL_PATH"] = str(COPIA_TEMPORAL)

import logging                                                  # noqa: E402

from conexion_drive import editar_excel, leer_excel             # noqa: E402
from conexion_drive.recalculo import buscar_errores             # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla     # noqa: E402
from . import activacion_promocion as ap                        # noqa: E402
from . import nodo_integral as ni                               # noqa: E402

logging.basicConfig(level=logging.WARNING)
ERRORES = []
HOJA_COMBOS = "Tabla Combos Dinamicos"


def paso(titulo):
    print(f"\n--- {titulo} ---")


def revisar(condicion, texto):
    print(f"  {'OK   ' if condicion else 'ERROR'} {texto}")
    if not condicion:
        ERRORES.append(texto)


def huella(ruta):
    return hashlib.md5(ruta.read_bytes()).hexdigest()


def ultimo_evento():
    ws, tabla = obtener_tabla(leer_excel(), "RegistroEventos", "tblRegistroEventos")
    return leer_tabla(ws, tabla)[-1]


def columna_c():
    """{celda: valor} de la columna C de NODO INTEGRAL, como la calcula el Excel."""
    ws = leer_excel()[ap.HOJA]
    return {f"C{f}": ws.cell(f, 3).value for f in range(3, ws.max_row + 1) if ws.cell(f, 2).value}


def _celdas_limite(ws):
    """Las celdas 'Valor limite' de los combos (se ubican por el encabezado, como hace catalogo.py)."""
    for fila in ws.iter_rows(min_row=1, max_row=15):
        textos = {str(c.value).strip(): c.column for c in fila if c.value is not None}
        if "Nombre del Combo (Key)" in textos:
            celdas = []
            for f in range(fila[0].row + 1, ws.max_row + 1):
                if ws.cell(f, textos["Nombre del Combo (Key)"]).value and ws.cell(f, textos["codigo"]).value:
                    celdas.append(ws.cell(f, textos["Valor límite"]))
            return celdas
    raise KeyError("No encontre los encabezados de Tabla Combos Dinamicos")


def simular(dia_pago, limites=None):
    """Escribe en la copia TEMPORAL el dia de pago (SI / NO) y, si se piden, los limites de los combos."""
    def _editar(wb, wb_valores):
        ws = wb[ap.HOJA]
        celdas = ni._celdas(ws)
        ws[celdas["dia_pago_lista"]] = dia_pago
        ws[celdas["dia_pago_hoy"]] = dia_pago
        if limites is not None:
            for celda, valor in zip(_celdas_limite(wb[HOJA_COMBOS]), limites):
                celda.value = valor
        return {"hubo_cambios": True}
    editar_excel(_editar, descripcion="Simulacion de la prueba")


def main():
    huella_local = huella(COPIA_LOCAL)
    print(f"Copia temporal de la prueba: {COPIA_TEMPORAL}")
    wb = leer_excel()
    celdas_nodo = ni._celdas(wb[ap.HOJA])
    temperatura = wb[ap.HOJA][celdas_nodo["temperatura"]].value
    dia_pago_inicial = wb[ap.HOJA][celdas_nodo["dia_pago_hoy"]].value
    limites_iniciales = [c.value for c in _celdas_limite(wb[HOJA_COMBOS])]
    inicio = columna_c()

    # La prueba parte de un dia normal (sin dia de pago ni feriado) con sus combos.
    if dia_pago_inicial != "NO" or wb[ap.HOJA][celdas_nodo["feriado_hoy"]].value != "NO":
        print("\nLa copia local esta en dia de pago o feriado: la prueba necesita un dia normal. No se puede seguir.")
        return 1

    paso("1) Dia normal con combos activos")
    antes = ap.leer_activacion()
    print(f"         C16 dice {antes['valor_celda']!r}; temperatura {temperatura}; corresponde {antes['valor']} ({antes['motivo']})")
    revisar(antes["celda"] == "C16", "la celda esta donde dice el mapa: C16")
    revisar(antes["valor"] == ap.ON and antes["tipo"] == ap.COMBOS and antes["combos"],
            f"corresponde ON por combos: {[c['codigo'] for c in antes['combos']]}")
    resultado = ap.actualizar_activacion_promocion(origen="Prueba")
    despues = columna_c()
    if antes["al_dia"]:
        revisar(resultado["estado"] == "SIN_CAMBIO", "C16 ya decia ON: SIN_CAMBIO")
    else:
        revisar(resultado["estado"] == "OK" and despues["C16"] == ap.ON, f"C16: {antes['valor_celda']!r} -> 'ON'")
        evento = ultimo_evento()
        revisar(evento.get("TIPO_EVENTO") == ap.TIPO_EVENTO and evento.get("ORIGEN") == "Prueba"
                and evento.get("ID_ENTIDAD") == "C16" and "combos activos" in str(evento.get("DETALLE")),
                f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    revisar(all(despues[c] == inicio[c] for c in inicio if c != "C16"), "las demas filas de la hoja quedaron iguales")

    paso("2) Otra vez, sin que nada cambie")
    huella_antes = huella(COPIA_TEMPORAL)
    resultado = ap.actualizar_activacion_promocion()
    revisar(resultado["estado"] == "SIN_CAMBIO" and huella(COPIA_TEMPORAL) == huella_antes, "SIN_CAMBIO: no se escribe nada")
    revisar(ap.leer_activacion()["al_dia"], "la lectura dice que C16 esta al dia")

    paso("3) Dia de pago: descuento general y sin combos (Regla #8)")
    simular("SI")
    estado = ap.leer_activacion()
    revisar(estado["valor"] == ap.ON and estado["tipo"] == ap.DESCUENTO_GENERAL and not estado["combos"]
            and estado["detonador"] == "DIA_PAGO", f"corresponde ON por {estado['motivo']}")
    huella_antes = huella(COPIA_TEMPORAL)
    resultado = ap.actualizar_activacion_promocion()
    revisar(resultado["estado"] == "SIN_CAMBIO" and huella(COPIA_TEMPORAL) == huella_antes,
            "C16 ya decia ON: SIN_CAMBIO, no se escribe nada")

    paso("4) Dia normal SIN ningun combo activo")
    # Limites que ningun combo cumple con la temperatura de la copia: '<=' muy bajo y '>=' muy alto.
    ws_combos = leer_excel()[HOJA_COMBOS]
    operadores = [ws_combos.cell(c.row, c.column - 1).value for c in _celdas_limite(ws_combos)]
    imposibles = [-99 if str(op).strip().startswith("<") else 99 for op in operadores]
    simular("NO", imposibles)
    estado = ap.leer_activacion()
    revisar(estado["valor"] == ap.OFF and estado["tipo"] == ap.NINGUNA and not estado["al_dia"],
            f"corresponde OFF ({estado['motivo']}) y C16 todavia dice {estado['valor_celda']!r}")
    resultado = ap.actualizar_activacion_promocion(origen="Prueba")
    evento = ultimo_evento()
    revisar(resultado["estado"] == "OK" and columna_c()["C16"] == ap.OFF, "C16: 'ON' -> 'OFF'")
    revisar("'ON' -> 'OFF'" in str(evento.get("DETALLE")), f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")

    paso("5) Vuelven los combos")
    simular("NO", limites_iniciales)
    resultado = ap.actualizar_activacion_promocion(origen="Prueba")
    final = columna_c()
    revisar(resultado["estado"] == "OK" and final["C16"] == ap.ON and resultado["tipo"] == ap.COMBOS, "C16: 'OFF' -> 'ON'")

    paso("6) Cierre")
    revisar(all(final[c] == inicio[c] for c in inicio if c != "C16"),
            "la hoja NODO INTEGRAL quedo igual que al inicio, menos C16 que ahora dice ON")
    ws = leer_excel(solo_valores=False)[ap.HOJA]
    formulas = [getattr(ws[c].value, "text", ws[c].value) for c in ("C19", "C20")]
    revisar(all(isinstance(f, str) and f.startswith("=_xlfn.TEXTJOIN") for f in formulas),
            "C19 y C20 siguen siendo formulas con _xlfn.TEXTJOIN (Regla #12)")
    revisar(not buscar_errores(COPIA_TEMPORAL.read_bytes()), "la copia temporal quedo con 0 errores de formula")
    revisar(huella(COPIA_LOCAL) == huella_local, "datos_prueba\\maestro_prueba.xlsx quedo IGUAL (no se toco)")

    print()
    if ERRORES:
        print(f"RESULTADO: {len(ERRORES)} linea(s) ERROR")
        return 1
    print("RESULTADO: TODO OK (0 lineas ERROR)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
