"""Prueba de existencia_minima.py SOBRE UNA COPIA TEMPORAL (no toca Drive ni la copia local).

Copia datos_prueba/maestro_prueba.xlsx a una carpeta temporal y trabaja ahi; al terminar
la carpeta se borra sola. Por eso NO hace falta "foto antes -> prueba -> restaurar".

Que revisa (Decisiones Pendientes #63, 'Control de Inventario' columna G):
  1) Que valores se aceptan: de 0 en adelante, hasta 2 decimales; lo demas se rechaza.
  2) Lectura de la hoja: todos los productos con su existencia, minimo y estado.
  3) Rechazos: producto que no esta en la hoja y valores invalidos NO escriben nada.
  4) Bajar el minimo de un producto en REABASTECER: solo cambia esa celda, el Excel pasa el
     estado a OK con su formula, y queda el evento con el valor anterior y el nuevo.
  5) El mismo valor otra vez: SIN_CAMBIO, no se escribe nada.
  6) Subir el minimo de un producto en OK: el Excel lo pasa a REABASTECER.
  7) Devolver los dos minimos: la hoja queda como al inicio.
  8) La columna Estado sigue siendo formula, 0 errores de formula y la copia local intacta.

Necesita LibreOffice (igual que las otras pruebas). Escribe 4 veces: tarda cerca de 1 minuto.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo2_inventario.probar_existencia_minima
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
_carpeta = Path(tempfile.mkdtemp(prefix="prueba_minimo_"))
COPIA_TEMPORAL = _carpeta / "copia_prueba.xlsx"
shutil.copy2(COPIA_LOCAL, COPIA_TEMPORAL)
atexit.register(shutil.rmtree, _carpeta, ignore_errors=True)

# Modo local sobre la copia temporal, ANTES de importar el robot (asi nunca toca Drive).
os.environ["ORIGEN_EXCEL"] = "local"
os.environ["EXCEL_LOCAL_PATH"] = str(COPIA_TEMPORAL)

import logging                                                  # noqa: E402

from conexion_drive import leer_excel                           # noqa: E402
from conexion_drive.recalculo import buscar_errores             # noqa: E402
from conexion_drive.tablas import leer_tabla, obtener_tabla     # noqa: E402
from . import existencia_minima as em                           # noqa: E402

logging.basicConfig(level=logging.WARNING)
ERRORES = []


def paso(titulo):
    print(f"\n--- {titulo} ---")


def revisar(condicion, texto):
    print(f"  {'OK   ' if condicion else 'ERROR'} {texto}")
    if not condicion:
        ERRORES.append(texto)


def huella(ruta):
    return hashlib.md5(ruta.read_bytes()).hexdigest()


def rechazado(funcion, tipo=em.MinimoInvalido):
    try:
        funcion()
    except tipo:
        return True
    return False


def foto():
    """{producto: (existencia, minimo, estado)} con todo redondeado a 2 decimales."""
    return {f["producto"]: (round(f["existencia"], 2), f["minimo"], f["estado"]) for f in em.leer_existencias()}


def ultimo_evento():
    ws, tabla = obtener_tabla(leer_excel(), "RegistroEventos", "tblRegistroEventos")
    return leer_tabla(ws, tabla)[-1]


def main():
    huella_local = huella(COPIA_LOCAL)
    print(f"Copia temporal de la prueba: {COPIA_TEMPORAL}")

    paso("1) Valores que se aceptan")
    revisar(em.normalizar_minimo("2") == 2 and em.normalizar_minimo(10) == 10, "'2' y 10 se aceptan como enteros")
    revisar(em.normalizar_minimo("0,5") == 0.5 and em.normalizar_minimo("2.50") == 2.5, "'0,5' y '2.50' se aceptan (coma o punto)")
    revisar(em.normalizar_minimo("1.239") == 1.24, "mas de 2 decimales: se redondea a 2 (1.239 -> 1.24)")
    revisar(em.normalizar_minimo("0") == 0, "0 se acepta (producto sin existencia minima)")
    for malo in ("-1", "abc", "", None, True, "nan", "inf", str(em.MAXIMO + 1)):
        revisar(rechazado(lambda malo=malo: em.normalizar_minimo(malo)), f"{malo!r}: rechazado")

    paso("2) Lectura de 'Control de Inventario'")
    inicio = foto()
    filas = em.leer_existencias()
    revisar(len(filas) >= 2 and all(isinstance(f["minimo"], (int, float)) for f in filas),
            f"{len(filas)} productos leidos, todos con existencia minima numerica")
    revisar(all(f["estado"] in ("OK", "REABASTECER") for f in filas), "el estado de cada producto es OK o REABASTECER")
    for f in filas:
        print(f"         fila {f['fila']:>2}  {f['producto'][:42]:<42} existencia {round(f['existencia'], 2):>8g}  "
              f"minimo {f['minimo']:>5g}  {f['estado']}")
    bajo = next((f for f in filas if f["estado"] == "REABASTECER" and f["existencia"] >= 0.1), None)
    alto = next((f for f in filas if f["estado"] == "OK"), None)
    if bajo is None or alto is None:
        print("\nLa copia local no tiene un producto en REABASTECER con existencia y otro en OK. No se puede seguir.")
        return 1

    paso("3) Rechazos: no se escribe nada")
    antes = huella(COPIA_TEMPORAL)
    revisar(rechazado(lambda: em.actualizar_existencia_minima("PRODUCTO QUE NO EXISTE", 5), em.ProductoNoEstaEnInventario),
            "un producto que no esta en la hoja: rechazado")
    revisar(rechazado(lambda: em.actualizar_existencia_minima(bajo["producto"].lower() + " x", 5)),
            "un nombre parecido pero no exacto: rechazado (los nombres son llave)")
    revisar(rechazado(lambda: em.actualizar_existencia_minima(bajo["producto"], -3)), "minimo negativo: rechazado")
    revisar(rechazado(lambda: em.actualizar_existencia_minima(bajo["producto"], "mucho")), "texto en vez de numero: rechazado")
    revisar(huella(COPIA_TEMPORAL) == antes, "el Excel no cambio")

    paso(f"4) Bajar el minimo de '{bajo['producto']}' (hoy REABASTECER)")
    nuevo_bajo = max(round(bajo["existencia"] / 2, 2), 0.01)
    resultado = em.actualizar_existencia_minima(bajo["producto"], nuevo_bajo, origen="Prueba")
    despues = foto()
    revisar(resultado["estado"] == "OK" and resultado["antes"] == bajo["minimo"] and resultado["despues"] == nuevo_bajo,
            f"celda {resultado['celda']}: {bajo['minimo']:g} -> {nuevo_bajo:g}")
    revisar(despues[bajo["producto"]][1] == nuevo_bajo, "el minimo nuevo quedo en la hoja")
    revisar(despues[bajo["producto"]][2] == "OK",
            f"el Excel recalculo el estado: existencia {round(bajo['existencia'], 2):g} ya no es menor que {nuevo_bajo:g} -> OK")
    revisar(despues[bajo["producto"]][0] == inicio[bajo["producto"]][0], "la existencia del producto no cambio")
    revisar(all(despues[p] == inicio[p] for p in inicio if p != bajo["producto"]), "los demas productos quedaron iguales")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == em.TIPO_EVENTO and evento.get("ORIGEN") == "Prueba" and "->" in str(evento.get("DETALLE"))
            and evento.get("ID_ENTIDAD") == resultado["celda"], f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")

    paso("5) El mismo valor otra vez")
    antes = huella(COPIA_TEMPORAL)
    resultado = em.actualizar_existencia_minima(bajo["producto"], nuevo_bajo)
    revisar(resultado["estado"] == "SIN_CAMBIO" and huella(COPIA_TEMPORAL) == antes, "SIN_CAMBIO: no se escribe nada")

    paso(f"6) Subir el minimo de '{alto['producto']}' (hoy OK)")
    nuevo_alto = int(alto["existencia"]) + 5
    em.actualizar_existencia_minima(alto["producto"], str(nuevo_alto), origen="Prueba")
    despues = foto()
    revisar(despues[alto["producto"]][1] == nuevo_alto and despues[alto["producto"]][2] == "REABASTECER",
            f"minimo {alto['minimo']:g} -> {nuevo_alto:g}: el Excel lo paso a REABASTECER")

    paso("7) Devolver los dos minimos a su valor")
    em.actualizar_existencia_minima(bajo["producto"], bajo["minimo"], origen="Prueba")
    em.actualizar_existencia_minima(alto["producto"], alto["minimo"], origen="Prueba")
    revisar(foto() == inicio, "'Control de Inventario' quedo igual que al inicio (existencias, minimos y estados)")

    paso("8) Cierre")
    ws, tabla = obtener_tabla(leer_excel(solo_valores=False), em.HOJA, em.TABLA)
    formulas = [f for f in leer_tabla(ws, tabla)]
    revisar(all(str(f.get(em.COL_ESTADO)).startswith("=") and str(f.get(em.COL_EXISTENCIA)).startswith("=") for f in formulas),
            "Existencia Actual y Estado siguen siendo formulas en todas las filas")
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
