"""Prueba de nodo_manual.py SOBRE UNA COPIA TEMPORAL (no toca Drive ni la copia local).

Copia datos_prueba/maestro_prueba.xlsx a una carpeta temporal y trabaja ahi; al terminar
la carpeta se borra sola. Por eso NO hace falta "foto antes -> prueba -> restaurar".

Que revisa (Decisiones Pendientes #64 y #80, hoja 'NODO INTEGRAL'):
  1) Que valores se aceptan en cada celda manual y cuales se rechazan.
  2) Lectura de la hoja: las 4 celdas manuales estan donde dice el mapa (C3, C4, C8, C22) y
     las 2 que quedaron solo lectura (C14, C15) no se pueden cambiar.
  3) Rechazos: no escriben nada.
  4) Tarifa Base KM (C3): cambia solo esa celda; 'Interface digital' B4 y el bot leen el
     valor nuevo; queda el evento con el valor anterior y el nuevo.
  5) El mismo valor otra vez: SIN_CAMBIO, no se escribe nada.
  6) Numero SINPE Movil (C22): se guarda con el formato 8888-8888 y el bot lee el nuevo.
  7) Meta de Venta Diaria (C4) y Accion Meta Ads (C8).
  8) Devolver C3, C4 y C22: las demas filas de la hoja nunca cambiaron.
  9) C19 y C20 siguen siendo formulas, 0 errores de formula y la copia local intacta.

Necesita LibreOffice (igual que las otras pruebas). Escribe 7 veces: tarda cerca de 1 minuto.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.probar_nodo_manual
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
_carpeta = Path(tempfile.mkdtemp(prefix="prueba_nodo_manual_"))
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
from microservicios.grupo1_captura_operativa.catalogo import leer_reglas_del_dia   # noqa: E402
from . import nodo_manual as nm                                 # noqa: E402

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


def rechazado(funcion, tipo=nm.ValorInvalido):
    try:
        funcion()
    except tipo:
        return True
    return False


def foto():
    """{celda: valor} de toda la columna C de NODO INTEGRAL, como la calcula el Excel."""
    return {f["celda"]: f["valor"] for f in nm.leer_nodo()}


def ultimo_evento():
    ws, tabla = obtener_tabla(leer_excel(), "RegistroEventos", "tblRegistroEventos")
    return leer_tabla(ws, tabla)[-1]


def main():
    huella_local = huella(COPIA_LOCAL)
    print(f"Copia temporal de la prueba: {COPIA_TEMPORAL}")

    paso("1) Valores que se aceptan y que se rechazan")
    revisar(nm.normalizar_valor("tarifa_km", "550") == 550 and nm.normalizar_valor("tarifa_km", 5000) == 5000
            and nm.normalizar_valor("tarifa_km", "500,00") == 500, "Tarifa Base KM: '550', 5000 y '500,00' se aceptan")
    for malo in ("0", "-5", "5001", "500.5", "1,500", "abc", "", None, True, "nan"):
        revisar(rechazado(lambda malo=malo: nm.normalizar_valor("tarifa_km", malo)), f"Tarifa Base KM {malo!r}: rechazado")
    revisar(nm.normalizar_valor("meta_venta", "0") == 0 and nm.normalizar_valor("meta_venta", "175000") == 175000
            and nm.normalizar_valor("meta_venta", "175000,50") == 175000.5, "Meta de Venta Diaria: '0', '175000' y '175000,50' se aceptan")
    for malo in ("-1", "150,000", "1.500.000", "100000000", "mucho", ""):
        revisar(rechazado(lambda malo=malo: nm.normalizar_valor("meta_venta", malo)), f"Meta de Venta Diaria {malo!r}: rechazado")
    revisar(nm.normalizar_valor("meta_ads", "subir") == "Subir" and nm.normalizar_valor("meta_ads", " Mantener ") == "Mantener",
            "Accion Meta Ads: 'subir' y ' Mantener ' se aceptan (se guardan Subir / Mantener)")
    for malo in ("Bajar", "[Subir / Mantener]", "", None):
        revisar(rechazado(lambda malo=malo: nm.normalizar_valor("meta_ads", malo)), f"Accion Meta Ads {malo!r}: rechazado")
    revisar(nm.normalizar_valor("sinpe", "70001234") == "7000-1234" and nm.normalizar_valor("sinpe", "7000 1234") == "7000-1234"
            and nm.normalizar_valor("sinpe", "8888-8888") == "8888-8888", "SINPE: '70001234', '7000 1234' y '8888-8888' se guardan con guion (7000-1234)")
    for malo in ("1234", "123456789", "7000-12a4", "=70001234", "+50670001234", ""):
        revisar(rechazado(lambda malo=malo: nm.normalizar_valor("sinpe", malo)), f"SINPE {malo!r}: rechazado")

    paso("2) Lectura de 'NODO INTEGRAL'")
    filas = nm.leer_nodo()
    inicio = foto()
    for f in filas:
        marca = {"manual": "MANUAL    ", "sin_efecto": "SIN EFECTO", "lectura": "lectura   "}[f["clase"]]
        print(f"         {f['celda']:>4}  {marca}  {f['descripcion'][:32]:<32} {str(f['valor'])[:44]}")
    manuales = {f["campo"]: f for f in filas if f["clase"] == "manual"}
    revisar(set(manuales) == set(nm.CAMPOS), f"las {len(nm.CAMPOS)} celdas manuales estan en la hoja")
    revisar(all(manuales[c]["celda"] == d["celda_mapa"] for c, d in nm.CAMPOS.items() if c in manuales),
            "estan donde dice el mapa: C3, C4, C8 y C22")
    sin_efecto = [f["celda"] for f in filas if f["clase"] == "sin_efecto"]
    revisar(sin_efecto == ["C14", "C15"], f"solo lectura por la Regla #8: {sin_efecto}")
    if set(manuales) != set(nm.CAMPOS):
        print("\nLa copia local no tiene las cuatro celdas manuales. No se puede seguir.")
        return 1
    tarifa_inicial, meta_inicial, sinpe_inicial = (manuales[c]["valor"] for c in ("tarifa_km", "meta_venta", "sinpe"))

    paso("3) Rechazos: no se escribe nada")
    antes = huella(COPIA_TEMPORAL)
    revisar(rechazado(lambda: nm.actualizar_dato_manual("oferta_feriado", "Combo 1"), nm.CampoDesconocido),
            "C14 (Oferta Fija por Feriado) no se puede cambiar: campo desconocido")
    revisar(rechazado(lambda: nm.actualizar_dato_manual("clima", "Lluvia"), nm.CampoDesconocido),
            "una fila que escribe el robot (clima) no se puede cambiar")
    revisar(rechazado(lambda: nm.actualizar_dato_manual("tarifa_km", 0)), "tarifa 0: rechazada")
    revisar(rechazado(lambda: nm.actualizar_dato_manual("sinpe", "1234")), "SINPE de 4 digitos: rechazado")
    revisar(rechazado(lambda: nm.actualizar_dato_manual("meta_ads", "Bajar")), "Accion Meta Ads 'Bajar': rechazada")
    revisar(huella(COPIA_TEMPORAL) == antes, "el Excel no cambio")

    paso("4) Tarifa Base KM (C3)")
    nueva_tarifa = 550 if tarifa_inicial != 550 else 600
    resultado = nm.actualizar_dato_manual("tarifa_km", str(nueva_tarifa), origen="Prueba")
    despues = foto()
    revisar(resultado["estado"] == "OK" and resultado["celda"] == "C3" and resultado["antes"] == tarifa_inicial
            and resultado["despues"] == nueva_tarifa, f"celda C3: {tarifa_inicial} -> {nueva_tarifa}")
    revisar(all(despues[c] == inicio[c] for c in inicio if c != "C3"), "las demas filas de la hoja quedaron iguales")
    wb = leer_excel()
    revisar(wb["Interface digital"]["B4"].value == nueva_tarifa, f"'Interface digital' B4 (formula) ya dice {nueva_tarifa}")
    revisar(leer_reglas_del_dia(wb)["tarifa_km"] == nueva_tarifa, "el bot (catalogo.py) lee la tarifa nueva para el envio")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == nm.TIPO_EVENTO and evento.get("ORIGEN") == "Prueba" and evento.get("ID_ENTIDAD") == "C3"
            and "->" in str(evento.get("DETALLE")), f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")

    paso("5) El mismo valor otra vez")
    antes = huella(COPIA_TEMPORAL)
    resultado = nm.actualizar_dato_manual("tarifa_km", nueva_tarifa)
    revisar(resultado["estado"] == "SIN_CAMBIO" and huella(COPIA_TEMPORAL) == antes, "SIN_CAMBIO: no se escribe nada")

    paso("6) Numero SINPE Movil (C22)")
    nuevo_sinpe = "7000-1234" if sinpe_inicial != "7000-1234" else "7000-4321"
    resultado = nm.actualizar_dato_manual("sinpe", nuevo_sinpe.replace("-", ""), origen="Prueba")
    wb = leer_excel()
    revisar(resultado["celda"] == "C22" and wb[nm.HOJA]["C22"].value == nuevo_sinpe,
            f"celda C22: {sinpe_inicial!r} -> {nuevo_sinpe!r} (digitado sin guion, guardado con guion)")
    revisar(leer_reglas_del_dia(wb)["sinpe_movil"] == nuevo_sinpe, "el bot (catalogo.py) lee el numero nuevo")

    paso("7) Meta de Venta Diaria (C4) y Accion Meta Ads (C8)")
    nueva_meta = 175000 if meta_inicial != 175000 else 180000
    resultado = nm.actualizar_dato_manual("meta_venta", nueva_meta, origen="Prueba")
    revisar(resultado["celda"] == "C4" and foto()["C4"] == nueva_meta, f"celda C4: {meta_inicial} -> {nueva_meta}")
    ads_inicial = manuales["meta_ads"]["valor"]
    nueva_ads = "Mantener" if str(ads_inicial).strip() == "Subir" else "Subir"
    resultado = nm.actualizar_dato_manual("meta_ads", nueva_ads.lower(), origen="Prueba")
    evento = ultimo_evento()
    revisar(resultado["celda"] == "C8" and foto()["C8"] == nueva_ads, f"celda C8: {ads_inicial!r} -> {nueva_ads!r}")
    revisar(repr(ads_inicial) in str(evento.get("DETALLE")), f"el evento guarda el texto anterior: {evento.get('DETALLE')}")
    fila_ads = next(f for f in nm.leer_nodo() if f["campo"] == "meta_ads")
    revisar(fila_ads["definido"], "la celda C8 ya tiene una opcion de la lista")

    paso("8) Devolver C3, C4 y C22 a su valor")
    nm.actualizar_dato_manual("tarifa_km", tarifa_inicial, origen="Prueba")
    nm.actualizar_dato_manual("meta_venta", meta_inicial, origen="Prueba")
    nm.actualizar_dato_manual("sinpe", sinpe_inicial, origen="Prueba")
    final = foto()
    revisar(all(final[c] == inicio[c] for c in inicio if c != "C8"),
            "la hoja quedo igual que al inicio (menos C8, que ahora tiene una opcion de la lista)")
    revisar(leer_excel()["Interface digital"]["B4"].value == tarifa_inicial, f"'Interface digital' B4 volvio a {tarifa_inicial}")

    paso("9) Cierre")
    ws = leer_excel(solo_valores=False)[nm.HOJA]
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
