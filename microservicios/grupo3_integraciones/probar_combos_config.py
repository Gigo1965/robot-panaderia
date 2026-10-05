"""Prueba de combos_config.py SOBRE UNA COPIA TEMPORAL (no toca Drive ni la copia local).

Copia datos_prueba/maestro_prueba.xlsx a una carpeta temporal y trabaja ahi; al terminar
la carpeta se borra sola. Por eso NO hace falta "foto antes -> prueba -> restaurar".

Que revisa (Decisiones Pendientes #65 y #89, hoja 'Tabla Combos Dinamicos'):
  1) Que valores se aceptan en cada celda (G a L) y cuales se rechazan.
  2) El catalogo (#89): la lista de combos termina en la primera fila vacia; "Baguette" por
     NOMBRE ya no se confunde con una fila de la tabla de composicion.
  3) Lectura: los combos estan donde dice el mapa (filas 5 a 8, celdas G a L).
  4) Rechazos: no escriben nada.
  5) Descuento (L): cambia solo esa celda; el Excel recalcula el Precio de Oferta (E); el robot
     alinea el texto de la columna B; queda UN evento con el valor anterior y el nuevo; lo ya
     vendido y el inventario no cambian (Regla #40).
  6) Lo mismo otra vez: SIN_CAMBIO, no se escribe nada.
  7) Descuento de 1 %: se guarda 0.01 (nunca queda un 100 %).
  8) Detonador: sin ningun combo ACTIVO, 'Activacion de Promocion' (C16) pasa a OFF (#84).
  9) Detonador LLUVIA con el clima de la hoja: el combo se activa, C16 vuelve a ON y
     'NODO INTEGRAL' C19 / C20 (formulas) lo muestran.
 10) Devolver todo: G a L, el estado y el precio de cada combo quedan como al inicio.
 11) A y E siguen siendo formulas, 0 errores de formula y la copia local intacta.

Necesita LibreOffice (igual que las otras pruebas). Escribe unas 10 veces: tarda cerca de 1.5 minutos.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.probar_combos_config
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
_carpeta = Path(tempfile.mkdtemp(prefix="prueba_combos_config_"))
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
from microservicios.grupo1_captura_operativa.catalogo import Catalogo   # noqa: E402
from . import activacion_promocion as ap                        # noqa: E402
from . import combos_config as cc                               # noqa: E402

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


def rechazado(funcion, tipo=cc.ValorInvalido):
    try:
        funcion()
    except tipo:
        return True
    return False


def combos():
    return {c["codigo"]: c for c in cc.leer_combos()}


def guardar(combo, **cambios):
    """Guarda un combo con su configuracion actual, cambiando solo lo indicado."""
    datos = {clave: combo[clave] for clave in cc.CAMPOS}
    datos.update(cambios)
    return cc.actualizar_combo(combo["codigo"], origen="Prueba", **datos)


def ultimos_eventos(cuantos=2):
    ws, tabla = obtener_tabla(leer_excel(), "RegistroEventos", "tblRegistroEventos")
    return leer_tabla(ws, tabla)[-cuantos:]


def evento_de(id_evento):
    ws, tabla = obtener_tabla(leer_excel(), "RegistroEventos", "tblRegistroEventos")
    return next(e for e in leer_tabla(ws, tabla) if e.get("ID_EVENTO") == id_evento)


def ventas_del_combo(nombre):
    """Los importes ya guardados de las ventas digitales de ese combo (son numeros, no formulas)."""
    ws, tabla = obtener_tabla(leer_excel(), "Detalle comanda digital", "Tabla3")
    return [round(float(f.get("SUBTOTAL_LINEA") or 0), 2) for f in leer_tabla(ws, tabla)
            if str(f.get("PRODUCTO") or "").strip() == nombre]


def existencias():
    ws, tabla = obtener_tabla(leer_excel(), "Control de Inventario", "TablaControlInventario")
    return {str(f.get("Producto Maestro")).strip(): round(float(f.get("Existencia Actual") or 0), 2)
            for f in leer_tabla(ws, tabla) if f.get("Producto Maestro")}


def precio_de(nombre_producto):
    return next(p["precio"] for p in Catalogo(leer_excel()).productos if str(p["nombre"]).strip() == nombre_producto)


def main():
    huella_local = huella(COPIA_LOCAL)
    print(f"Copia temporal de la prueba: {COPIA_TEMPORAL}")

    paso("1) Valores que se aceptan y que se rechazan")
    n = cc.normalizar_configuracion
    bueno = n("temperatura", "< =", "27,5", "3", "  Hola   combo ", "10")
    revisar(bueno == {"tipo": "TEMPERATURA", "operador": "<=", "valor": 27.5, "prioridad": 3, "mensaje": "Hola combo",
                      "descuento": 10},
            "TEMPERATURA '< =' 27,5 prioridad 3 descuento 10 -> se guarda '<=', 27.5, 3 y 10 (sin espacios de mas)")
    lluvia = n("Lluvia", "", "despejado", 9, "Combo de sol", "0", operador_actual=">=")
    revisar(lluvia["tipo"] == "LLUVIA" and lluvia["valor"] == "Despejado" and lluvia["operador"] == ">=",
            "LLUVIA con 'despejado' -> 'Despejado'; el operador no se usa y queda el que tenia la celda")
    revisar(cc.descuento_a_celda(10) == 10 and cc.descuento_a_celda(2.5) == 2.5 and cc.descuento_a_celda(1) == 0.01
            and cc.descuento_a_celda(0.5) == 0.005 and cc.descuento_a_celda(0) == 0,
            "descuento en la celda: 10 -> 10, 2.5 -> 2.5, 1 -> 0.01, 0.5 -> 0.005, 0 -> 0")
    revisar(cc.descuento_de_celda(10) == 10 and cc.descuento_de_celda(0.01) == 1 and cc.descuento_de_celda(0.1) == 10
            and cc.descuento_de_celda(None) == 0,
            "descuento leido de la celda: 10 -> 10 %, 0.01 -> 1 %, 0.1 -> 10 %, vacia -> 0 %")
    base = dict(tipo="TEMPERATURA", operador="<=", valor="30", prioridad="1", mensaje="Combo", descuento="10")
    malos = {
        "detonador FERIADO (Regla #8)": dict(tipo="FERIADO"), "detonador DIA_PAGO (Regla #8)": dict(tipo="DIA_PAGO"),
        "detonador inventado": dict(tipo="VIENTO"), "operador '<>'": dict(operador="<>"), "operador '='": dict(operador="="),
        "temperatura 51": dict(valor="51"), "temperatura -1": dict(valor="-1"), "temperatura 30.25": dict(valor="30.25"),
        "temperatura con texto": dict(valor="calor"), "LLUVIA con 'Nublado'": dict(tipo="LLUVIA", valor="Nublado"),
        "prioridad 0": dict(prioridad="0"), "prioridad 10": dict(prioridad="10"), "prioridad 1.5": dict(prioridad="1.5"),
        "mensaje vacio": dict(mensaje="   "), "mensaje de 151 caracteres": dict(mensaje="x" * 151),
        "mensaje que empieza con '='": dict(mensaje="=1+1"), "descuento 11": dict(descuento="11"),
        "descuento -1": dict(descuento="-1"), "descuento 5.55": dict(descuento="5.55"),
        "descuento con texto": dict(descuento="diez"),
    }
    for nombre, cambio in malos.items():
        revisar(rechazado(lambda c=cambio: n(**dict(base, **c))), f"{nombre}: rechazado")
    revisar(len(n(**dict(base, mensaje="x" * 150))["mensaje"]) == 150, "mensaje de 150 caracteres: aceptado")

    paso("2) El catalogo lee solo la lista de combos (#89)")
    catalogo = Catalogo(leer_excel())
    inicio = combos()
    revisar([c["codigo"] for c in catalogo.combos] == list(inicio) and len(inicio) >= 1,
            f"el catalogo tiene {len(catalogo.combos)} combos ({', '.join(inicio)}): ninguna fila de la tabla de composicion")
    por_nombre, por_codigo = catalogo.resolver("Baguette"), catalogo.resolver("P01")
    revisar(por_nombre["origen"] == "TABLA CRUCE MAESTRA" and por_nombre["precio"] == por_codigo["precio"],
            f"'Baguette' por nombre y 'P01' por codigo dan el mismo producto y precio ({por_nombre['precio']:,.2f})")
    revisar(catalogo.resolver("CupCake")["origen"] == "TABLA CRUCE MAESTRA", "'CupCake' por nombre: producto de la tabla maestra")

    paso("3) Lectura de 'Tabla Combos Dinamicos'")
    primero = next(iter(inicio.values()))
    revisar([c["fila"] for c in inicio.values()] == list(range(primero["fila"], primero["fila"] + len(inicio))),
            f"{len(inicio)} combos seguidos, filas {primero['fila']} a {primero['fila'] + len(inicio) - 1}")
    revisar([primero["celdas"][c][0] for c in ("clima",) + cc.CAMPOS] == list("BGHIJKL"),
            f"columnas segun el mapa: B criterio, G tipo, H operador, I valor, J prioridad, K mensaje, L descuento "
            f"({primero['celdas']['tipo']} a {primero['celdas']['descuento']})")
    revisar(all(c["configurable"] for c in inicio.values()), "todos los combos usan TEMPERATURA o LLUVIA")
    for c in inicio.values():
        print(f"       {c['codigo']} {c['nombre']}: {'ACTIVO' if c['activo'] else 'INACTIVO'}, {c['tipo']} {c['operador']} "
              f"{c['valor']}, prioridad {c['prioridad']}, descuento {c['descuento']:g} %, precio {c['precio']:,.2f}")

    paso("4) Rechazos: no se escribe nada")
    huella_antes = huella(COPIA_TEMPORAL)
    revisar(rechazado(lambda: guardar(dict(primero, codigo="co99")), cc.ComboDesconocido), "combo 'co99': no existe, rechazado")
    revisar(rechazado(lambda: guardar(dict(primero, codigo="Baguette")), cc.ComboDesconocido),
            "'Baguette' (una fila de la tabla de composicion) no es un combo: rechazado")
    revisar(rechazado(lambda: guardar(primero, tipo="FERIADO")) and rechazado(lambda: guardar(primero, descuento="30"))
            and rechazado(lambda: guardar(primero, mensaje="=SUMA(A1)")) and rechazado(lambda: guardar(primero, valor="80")),
            "detonador FERIADO, descuento 30, mensaje con '=' y temperatura 80: rechazados")
    revisar(huella(COPIA_TEMPORAL) == huella_antes, "la copia temporal no cambio con los rechazos")

    paso("5) Descuento del primer combo")
    combo = primero
    partes = None
    ws_formulas = leer_excel(solo_valores=False)[cc.HOJA]
    formula_precio = str(ws_formulas[combo["celdas"]["precio"]].value)
    ventas_antes, existencias_antes = ventas_del_combo(combo["nombre"]), existencias()
    descuento_nuevo = 5 if combo["descuento"] != 5 else 7.5
    resultado = guardar(combo, descuento=str(descuento_nuevo))
    ahora = combos()[combo["codigo"]]
    celdas_cambiadas = sorted(c["celda"] for c in resultado["cambios"])
    revisar(resultado["estado"] == "OK" and combo["celdas"]["descuento"] in celdas_cambiadas
            and set(celdas_cambiadas) <= {combo["celdas"]["descuento"], combo["celdas"]["clima"]},
            f"{combo['nombre']}: descuento {combo['descuento']:g} % -> {descuento_nuevo:g} %; celdas escritas: {', '.join(celdas_cambiadas)}")
    revisar(ahora["descuento"] == descuento_nuevo and all(ahora[c] == combo[c] for c in cc.CAMPOS if c != "descuento"),
            "las otras cinco celdas del combo (G a K) quedaron igual")
    revisar(ahora["clima"] == cc.texto_criterio(ahora["tipo"], ahora["operador"], ahora["valor"]),
            f"el robot dejo el texto de la columna B igual al criterio: {combo['clima']!r} -> {ahora['clima']!r}")
    # El precio lo calcula el Excel: parte con descuento (produccion propia) + parte sin descuento (reventa).
    factor_antes, factor_nuevo = 1 - combo["descuento"] / 100, 1 - descuento_nuevo / 100
    if factor_antes > 0 and combo["precio"] is not None:
        # precio = propio * factor + reventa  ->  con dos precios se despejan las dos partes
        propio = (combo["precio"] - ahora["precio"]) / (factor_antes - factor_nuevo)
        reventa = combo["precio"] - propio * factor_antes
        partes = (propio, reventa)
        revisar(propio > 0 and reventa >= 0 and round(propio * factor_nuevo + reventa, 2) == round(ahora["precio"], 2),
                f"el Excel recalculo el Precio de Oferta: {combo['precio']:,.2f} -> {ahora['precio']:,.2f} "
                f"(el descuento se aplica solo a la parte de produccion propia, {propio:,.2f}; Regla #28)")
    revisar(str(leer_excel(solo_valores=False)[cc.HOJA][combo["celdas"]["precio"]].value) == formula_precio,
            "la formula del precio (columna E) no se toco")
    evento = evento_de(resultado["evento"])
    revisar(evento.get("TIPO_EVENTO") == cc.TIPO_EVENTO and evento.get("ORIGEN") == "Prueba"
            and evento.get("ID_ENTIDAD") == combo["codigo"] and "->" in str(evento.get("DETALLE")),
            f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    revisar(ventas_del_combo(combo["nombre"]) == ventas_antes and existencias() == existencias_antes,
            f"lo ya vendido no cambio: {len(ventas_antes)} venta(s) digitales de {combo['nombre']} con el mismo importe, "
            "y el inventario igual (Regla #40)")
    activacion = ap.leer_activacion()
    revisar(resultado["activacion"]["valor"] == activacion["valor"] and activacion["al_dia"],
            f"despues de guardar, el robot reviso C16: {activacion['valor_celda']} ({activacion['motivo']})")

    paso("6) Lo mismo otra vez")
    huella_antes = huella(COPIA_TEMPORAL)
    resultado = guardar(combos()[combo["codigo"]])
    revisar(resultado["estado"] == "SIN_CAMBIO" and huella(COPIA_TEMPORAL) == huella_antes,
            "sin cambios en G a L: SIN_CAMBIO, no se escribe nada")

    paso("7) Descuento de 1 %")
    guardar(combos()[combo["codigo"]], descuento="1")
    ahora = combos()[combo["codigo"]]
    celda_l = leer_excel()[cc.HOJA][combo["celdas"]["descuento"]].value
    revisar(celda_l == 0.01 and ahora["descuento"] == 1, f"la celda {combo['celdas']['descuento']} guarda 0.01 y se lee como 1 %")
    if partes:
        revisar(round(partes[0] * 0.99 + partes[1], 2) == round(ahora["precio"], 2),
                f"el precio baja solo 1 % en la parte propia: {ahora['precio']:,.2f} (no queda en 100 % de descuento)")

    paso("8) Sin combos activos: 'Activacion de Promocion' pasa a OFF")
    activos = [c for c in combos().values() if c["activo"]]
    revisar(len(activos) >= 1, f"combos ACTIVOS al inicio: {', '.join(c['codigo'] for c in activos)}")
    for c in activos:                                   # mayor que 50 grados: nunca se cumple (el tope es 50)
        resultado = guardar(c, tipo="TEMPERATURA", operador=">", valor="50")
    ahora = combos()
    revisar(not any(c["activo"] for c in ahora.values()), "detonador 'TEMPERATURA > 50' en los activos: los combos quedan INACTIVOS")
    activacion = ap.leer_activacion()
    reglas = Catalogo(leer_excel()).reglas
    if reglas["descuento_general"]:
        revisar(activacion["valor_celda"] == "ON", "la copia esta en feriado o dia de pago: C16 sigue en ON por el descuento general")
    else:
        revisar(activacion["valor_celda"] == "OFF" and activacion["al_dia"]
                and resultado["activacion"]["estado"] == "OK",
                f"C16 paso a OFF ({activacion['motivo']})")
        evento = ultimos_eventos(1)[0]
        revisar(evento.get("TIPO_EVENTO") == ap.TIPO_EVENTO and evento.get("ORIGEN") == "Prueba",
                f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")

    paso("9) Detonador LLUVIA con el clima de la hoja")
    clima_hoy = str(leer_excel()["NODO INTEGRAL"]["C5"].value or "").strip()
    elegido = activos[-1]
    if clima_hoy not in cc.CLIMAS:
        revisar(False, f"'NODO INTEGRAL' C5 dice {clima_hoy!r}: se esperaba Lluvia o Despejado")
    else:
        resultado = guardar(ahora[elegido["codigo"]], tipo="LLUVIA", valor=clima_hoy)
        ahora = combos()
        revisar(ahora[elegido["codigo"]]["activo"] and ahora[elegido["codigo"]]["clima"] == clima_hoy,
                f"{elegido['nombre']} con LLUVIA = {clima_hoy} (el clima de C5): ACTIVO; columna B dice {clima_hoy!r}")
        otro = "Despejado" if clima_hoy == "Lluvia" else "Lluvia"
        activacion = ap.leer_activacion()
        if not reglas["descuento_general"]:
            revisar(activacion["valor_celda"] == "ON" and [c["codigo"] for c in activacion["combos"]] == [elegido["codigo"]],
                    f"C16 volvio a ON ({activacion['motivo']})")
        nodo = leer_excel()["NODO INTEGRAL"]
        revisar(str(nodo["C20"].value).strip() == elegido["codigo"] and ahora[elegido["codigo"]]["mensaje"] in str(nodo["C19"].value),
                f"'NODO INTEGRAL' C20 (formula) dice {nodo['C20'].value!r} y C19 trae su mensaje")
        guardar(ahora[elegido["codigo"]], valor=otro)
        revisar(not combos()[elegido["codigo"]]["activo"], f"con LLUVIA = {otro}: INACTIVO")

    paso("10) Devolver la configuracion inicial")
    for codigo, original in inicio.items():
        actual = combos()[codigo]
        if any(actual[c] != original[c] for c in cc.CAMPOS):
            guardar(original)
    final = combos()
    revisar(all(final[k][c] == inicio[k][c] for k in inicio for c in cc.CAMPOS),
            "las celdas G a L de todos los combos quedaron como al inicio")
    revisar(all(final[k]["activo"] == inicio[k]["activo"] and round(final[k]["precio"], 2) == round(inicio[k]["precio"], 2)
                for k in inicio), "el estado (ACTIVO / INACTIVO) y el precio de cada combo son los del inicio")
    revisar(ventas_del_combo(combo["nombre"]) == ventas_antes and existencias() == existencias_antes,
            "las ventas ya hechas y el inventario siguen igual")
    activacion = ap.leer_activacion()
    revisar(activacion["al_dia"], f"C16 quedo al dia: {activacion['valor_celda']} ({activacion['motivo']})")

    paso("11) Cierre")
    ws = leer_excel(solo_valores=False)[cc.HOJA]
    formulas = [ws[c["celdas"][clave]].value for c in inicio.values() for clave in ("estado", "precio")]
    revisar(all(isinstance(f, str) and f.startswith("=") for f in formulas),
            "las columnas A (Detonador) y E (Precio de Oferta) siguen siendo formulas")
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
