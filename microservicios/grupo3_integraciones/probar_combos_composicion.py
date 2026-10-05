"""Prueba de combos_composicion.py SOBRE UNA COPIA TEMPORAL (no toca Drive ni la copia local).

Copia datos_prueba/maestro_prueba.xlsx a una carpeta temporal y trabaja ahi; al terminar
la carpeta se borra sola. Por eso NO hace falta "foto antes -> prueba -> restaurar".

Que revisa (Decisiones Pendientes #85, hoja 'Tabla Combos Dinamicos', TablaComposicionCombos):
  1) Que cantidades se aceptan (enteros de 1 a 10) y como se arma y se lee la formula del precio.
  2) Lectura: cada combo tiene un componente de produccion propia y uno de reventa, y la formula
     del precio dice lo mismo que la tabla de composicion.
  3) Rechazos: no escriben nada.
  4) Se vende el combo ANTES del cambio (para que haya algo que congelar).
  5) Cambia la CANTIDAD del componente propio: G = combos vendidos, H = unidades ya rebajadas,
     D = cantidad nueva; lo ya rebajado NO se mueve; el Excel recalcula el precio con la formula
     nueva; el detalle (D) y el mensaje (K) dicen la composicion nueva; UN evento; el precio de
     las ventas ya hechas no cambia (Regla #40).
  6) Lo mismo otra vez: SIN_CAMBIO, no se escribe nada.
  7) Una venta DESPUES del cambio rebaja con la cantidad nueva.
  8) Cambia el PRODUCTO de reventa: lo ya rebajado queda cargado al producto anterior en una fila
     de historia y la fila vigente empieza en 0 con el producto nuevo.
  9) Una venta despues rebaja el producto nuevo y no toca el anterior.
 10) Se vuelve al producto anterior y se cambia otra vez: la historia se ACUMULA en la fila que
     ya existe (la tabla no crece de nuevo).
 11) Devolver todo: la composicion y el precio del inicio; el inventario solo se movio por las
     ventas de la prueba, cada una con la composicion que regia cuando se hizo.
 12) Las formulas siguen siendo formulas, 0 errores de formula, los combos siguen siendo los
     mismos (las filas de historia no se confunden con combos) y la copia local queda intacta.

Si la copia esta en dia de pago o feriado (ese dia no se venden combos, Regla #8) la prueba lo
avisa y revisa lo mismo sin las ventas nuevas.

Necesita LibreOffice (igual que las otras pruebas). Escribe 9 veces: tarda cerca de 2.5 minutos.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.probar_combos_composicion
"""
import atexit
import hashlib
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent.parent
COPIA_LOCAL = RAIZ_PROYECTO / "datos_prueba" / "maestro_prueba.xlsx"
if not COPIA_LOCAL.exists():
    sys.exit(f"No encontre la copia local: {COPIA_LOCAL}\n"
             "La prueba necesita datos_prueba\\maestro_prueba.xlsx (la misma de las otras pruebas).")
_carpeta = Path(tempfile.mkdtemp(prefix="prueba_combos_composicion_"))
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
from microservicios.grupo1_captura_operativa.catalogo import Catalogo                    # noqa: E402
from microservicios.grupo1_captura_operativa.ventas_despacho import registrar_venta_despacho  # noqa: E402
from . import combos_composicion as co                          # noqa: E402
from . import combos_config as cc                               # noqa: E402

logging.basicConfig(level=logging.WARNING)
ERRORES = []
_ventas = [0]


def paso(titulo):
    print(f"\n--- {titulo} ---")


def revisar(condicion, texto):
    print(f"  {'OK   ' if condicion else 'ERROR'} {texto}")
    if not condicion:
        ERRORES.append(texto)


def huella(ruta):
    return hashlib.md5(ruta.read_bytes()).hexdigest()


def rechazado(funcion):
    try:
        funcion()
    except cc.ValorInvalido:
        return True
    return False


def composicion(codigo):
    return next(c for c in co.leer_composicion()["combos"] if c["codigo"] == codigo)


def inventario():
    """{producto: (Salidas en Combos, Existencia Actual)} de 'Control de Inventario' (lo calcula el Excel)."""
    ws, tabla = obtener_tabla(leer_excel(), "Control de Inventario", "TablaControlInventario")
    return {str(f.get("Producto Maestro")).strip(): (round(float(f.get("Salidas en Combos") or 0), 2),
                                                     round(float(f.get("Existencia Actual") or 0), 2))
            for f in leer_tabla(ws, tabla) if f.get("Producto Maestro")}


def precios_vendidos(nombre_combo):
    """Los precios ya guardados de las ventas de ese combo (despacho y digitales): son numeros."""
    wb = leer_excel()
    ws, tabla = obtener_tabla(wb, "Ventas despacho", "Tabla6")
    despacho = [round(float(f.get("Precio") or 0), 2) for f in leer_tabla(ws, tabla)
                if str(f.get("Producto") or "").strip() == nombre_combo]
    ws, tabla = obtener_tabla(wb, "Detalle comanda digital", "Tabla3")
    digitales = [round(float(f.get("SUBTOTAL_LINEA") or 0), 2) for f in leer_tabla(ws, tabla)
                 if str(f.get("PRODUCTO") or "").strip() == nombre_combo]
    return despacho, digitales


def tabla_composicion():
    """(rango de la tabla, numero de filas con datos)."""
    ws, tabla = obtener_tabla(leer_excel(solo_valores=False), co.HOJA, co.TABLA)
    return tabla.ref, len(leer_tabla(ws, tabla))


def ultimo_evento():
    ws, tabla = obtener_tabla(leer_excel(), "RegistroEventos", "tblRegistroEventos")
    return leer_tabla(ws, tabla)[-1]


def vender(codigo, cantidad):
    """Una venta de despacho del combo, con una hora distinta cada vez."""
    _ventas[0] += 1
    return registrar_venta_despacho(codigo, cantidad, momento=datetime(2026, 10, 5, 8, 0, _ventas[0]))


def main():
    huella_local = huella(COPIA_LOCAL)
    print(f"Copia temporal de la prueba: {COPIA_TEMPORAL}")

    paso("1) Cantidades y formula del precio")
    n = co.normalizar_cantidad
    revisar(n("3") == 3 and n(10) == 10 and n("1,0") == 1 and n(" 4 ") == 4,
            "cantidades '3', 10, '1,0' y ' 4 ' -> 3, 10, 1 y 4")
    revisar(all(rechazado(lambda v=v: n(v)) for v in ("0", "11", "2.5", "-1", "", "tres", True)),
            "cantidades 0, 11, 2.5, -1, vacia, 'tres' y VERDADERO: rechazadas")
    nueva = co.formula_precio(5, "L", "G", 4, 8, 2, 13)
    revisar(nueva == "=(4*'TABLA CRUCE MAESTRA'!G8)*(1-IF($L5>1,$L5/100,$L5))+2*'TABLA CRUCE MAESTRA'!G13"
            and co.leer_formula_precio(nueva, 5, "L", "G") == (4, 8, 2, 13),
            "la formula nueva se arma con el descuento solo sobre el componente propio y se vuelve a leer igual")
    vieja = "=(3*'TABLA CRUCE MAESTRA'!G8)*(1-IF($L5>1,$L5/100,$L5)) +'TABLA CRUCE MAESTRA'!G10"
    revisar(co.leer_formula_precio(vieja, 5, "L", "G") == (3, 8, 1, 10)
            and co.leer_formula_precio(vieja, 6, "L", "G") is None
            and co.leer_formula_precio("=3*5", 5, "L", "G") is None,
            "la formula que ya tenia el Excel (con espacio y sin el '1*') se lee como 3 y 1; una de otra fila o de otra forma, no")
    revisar(co.texto_detalle(3, "Baguette", 1.0, " Leche ") == "3 Baguette + 1 Leche",
            "el detalle se escribe '3 Baguette + 1 Leche' (cantidad, nombre de la TCM, signo +)")

    paso("2) Lectura de la composicion")
    leido = co.leer_composicion()
    combos_inicio = cc.leer_combos()
    revisar(len(leido["combos"]) == len(combos_inicio) >= 1 and all(c["ajustable"] for c in leido["combos"]),
            f"{len(leido['combos'])} combos, todos con un componente propio y uno de reventa vigentes")
    for combo in leido["combos"]:
        if combo["ajustable"]:
            print(f"       {combo['codigo']} {combo['nombre']}: "
                  + " + ".join(f"{c['cantidad']:g} {c['producto']} (fila {c['fila']}, vendidos {c['vendidos']:g}, "
                               f"unidades {c['unidades']:g})" for c in combo["componentes"]))
        else:
            print(f"       {combo['codigo']} {combo['nombre']}: NO ajustable: {combo['motivo']}")
    revisar(len(leido["propios"]) >= 1 and len(leido["terceros"]) >= 2
            and all(p["tipo"] == co.PROPIO for p in leido["propios"]) and all(p["tipo"] == co.TERCERO for p in leido["terceros"]),
            f"listas de la TCM: {len(leido['propios'])} de produccion propia y {len(leido['terceros'])} de reventa")
    hoja_formulas = leer_excel(solo_valores=False)[co.HOJA]
    coinciden = True
    for combo in leido["combos"]:
        if not combo["ajustable"]:
            continue
        propio, reventa = combo["componentes"]
        fila_tcm = {p["nombre"]: p["fila"] for p in leido["propios"] + leido["terceros"]}
        en_formula = co.leer_formula_precio(hoja_formulas[combo["celdas"]["precio"]].value, combo["fila"], "L", "G")
        coinciden = coinciden and en_formula == (propio["cantidad"], fila_tcm[propio["producto"]],
                                                 reventa["cantidad"], fila_tcm[reventa["producto"]])
    revisar(coinciden, "en cada combo la formula del precio usa las mismas cantidades y productos que la tabla de composicion")

    reglas = Catalogo(leer_excel()).reglas
    puede_vender = not reglas["descuento_general"]
    activos = [c for c in combos_inicio if c["activo"] and c["configurable"]]
    elegido = (activos or combos_inicio)[0]
    codigo, nombre = elegido["codigo"], elegido["nombre"]
    if not puede_vender:
        print("       (la copia esta en dia de pago o feriado: ese dia no se venden combos; se prueba sin ventas nuevas)")
    elif not elegido["activo"]:
        # Ningun combo ACTIVO en esta copia: se activa el primero para poder venderlo (se devuelve al final).
        cc.actualizar_combo(codigo, "TEMPERATURA", ">=", 0, elegido["prioridad"], elegido["mensaje"],
                            elegido["descuento"], origen="Prueba")
        print(f"       (ningun combo estaba ACTIVO: se activo {nombre} con TEMPERATURA >= 0 para la prueba)")
    inicio = composicion(codigo)
    propio, reventa = inicio["componentes"]
    otro = next(p for p in leido["terceros"] if p["nombre"] != reventa["producto"])
    print(f"       combo de la prueba: {codigo} {nombre}; producto de reventa alterno: {otro['nombre']}")

    def guardar(producto_1=None, cantidad_1=None, producto_2=None, cantidad_2=None):
        actual = composicion(codigo)["componentes"]
        return co.actualizar_composicion(
            codigo, producto_1 or actual[0]["producto"], cantidad_1 or actual[0]["cantidad"],
            producto_2 or actual[1]["producto"], cantidad_2 or actual[1]["cantidad"], origen="Prueba")

    paso("3) Rechazos (no se escribe nada)")
    huella_antes = huella(COPIA_TEMPORAL)
    casos = [
        ("un combo que no existe", lambda: co.actualizar_composicion("co99", propio["producto"], 3, reventa["producto"], 1)),
        ("componente 1 de reventa", lambda: guardar(producto_1=reventa["producto"])),
        ("componente 2 de produccion propia", lambda: guardar(producto_2=propio["producto"])),
        ("un producto que no esta en la TCM", lambda: guardar(producto_2="Pan de yuca")),
        ("producto vacio", lambda: co.actualizar_composicion(codigo, "", 3, reventa["producto"], 1)),
        ("cantidad 0", lambda: co.actualizar_composicion(codigo, propio["producto"], 0, reventa["producto"], 1)),
        ("cantidad 11", lambda: guardar(cantidad_1=11)),
        ("cantidad 2.5", lambda: guardar(cantidad_2="2.5")),
        ("cantidad 'tres'", lambda: guardar(cantidad_2="tres")),
    ]
    fallaron = [titulo for titulo, funcion in casos if not rechazado(funcion)]
    revisar(not fallaron and huella(COPIA_TEMPORAL) == huella_antes,
            f"{len(casos)} casos rechazados y el archivo no cambio: " + ", ".join(t for t, _ in casos)
            + (f" || NO se rechazaron: {fallaron}" if fallaron else ""))

    paso("4) Venta ANTES del cambio")
    esperado = {}          # producto -> Salidas en Combos que deberia mostrar el Excel

    def sumar(producto, unidades):
        esperado[producto] = round(esperado.get(producto, inventario_inicio.get(producto, (0, 0))[0]) + unidades, 2)

    def inventario_como_se_espera(texto):
        ahora = inventario()
        bien = all(ahora.get(p, (None,))[0] == salidas for p, salidas in esperado.items())
        bien = bien and all(round(inventario_inicio[p][1] - (salidas - inventario_inicio[p][0]), 2) == ahora[p][1]
                            for p, salidas in esperado.items() if p in inventario_inicio)
        revisar(bien, texto + ": " + "; ".join(f"{p} salidas en combos {s:g}" for p, s in esperado.items()))

    inventario_inicio = inventario()
    for producto in (propio["producto"], reventa["producto"], otro["nombre"]):
        sumar(producto, 0)
    if puede_vender:
        venta = vender(codigo, 2)
        sumar(propio["producto"], 2 * propio["cantidad"])
        sumar(reventa["producto"], 2 * reventa["cantidad"])
        revisar(venta["estado"] == "REGISTRADA", f"2 x {nombre} a {venta['precio']:,.2f} (fila {venta['fila']} de 'Ventas despacho')")
        inventario_como_se_espera("el Excel rebajo el inventario con la composicion del inicio")
    antes = composicion(codigo)
    propio, reventa = antes["componentes"]
    precios_antes = precios_vendidos(nombre)
    print(f"       {nombre} tiene {propio['vendidos']:g} vendidos: {propio['unidades']:g} {propio['producto']} y "
          f"{reventa['unidades']:g} {reventa['producto']} ya rebajados")

    paso("5) Cambia la CANTIDAD del componente propio")
    cantidad_nueva = propio["cantidad"] + 1 if propio["cantidad"] < co.CANTIDAD_MAXIMA else propio["cantidad"] - 1
    resultado = guardar(cantidad_1=cantidad_nueva)
    ahora = composicion(codigo)
    p, r = ahora["componentes"]
    revisar(resultado["estado"] == "OK" and p["cantidad"] == cantidad_nueva and p["congelados"] == propio["vendidos"]
            and p["rebajadas"] == propio["unidades"],
            f"{p['producto']} (fila {p['fila']}): cantidad {propio['cantidad']:g} -> {cantidad_nueva:g}; "
            f"G = {p['congelados']:g} combos congelados y H = {p['rebajadas']:g} unidades ya rebajadas")
    revisar(p["unidades"] == propio["unidades"] and (r["cantidad"], r["congelados"], r["rebajadas"], r["unidades"])
            == (reventa["cantidad"], reventa["congelados"], reventa["rebajadas"], reventa["unidades"]),
            f"lo ya rebajado no se movio ({p['unidades']:g} {p['producto']}) y la fila del otro componente quedo igual")
    inventario_como_se_espera("'Control de Inventario' quedo igual")
    precios = {x["nombre"]: x["precio"] for x in co.leer_composicion()["propios"] + co.leer_composicion()["terceros"]}
    descuento = elegido["descuento"] / 100
    precio_esperado = round(cantidad_nueva * precios[p["producto"]] * (1 - descuento) + r["cantidad"] * precios[r["producto"]], 2)
    revisar(round(ahora["precio"], 2) == precio_esperado and round(ahora["precio"], 2) != round(antes["precio"], 2),
            f"el Excel recalculo el precio con la formula nueva: {antes['precio']:,.2f} -> {ahora['precio']:,.2f} "
            f"(descuento de {elegido['descuento']:g} % solo sobre {p['producto']})")
    detalle_nuevo = co.texto_detalle(cantidad_nueva, p["producto"], r["cantidad"], r["producto"])
    revisar(ahora["detalle"] == detalle_nuevo, f"detalle (celda {ahora['celdas']['detalle']}): {ahora['detalle']}")
    if antes["detalle"] and antes["detalle"] in antes["mensaje"]:
        revisar(resultado["mensaje"] == "ACTUALIZADO" and detalle_nuevo in ahora["mensaje"] and antes["detalle"] not in ahora["mensaje"],
                f"el mensaje traia el detalle viejo y ahora dice: {ahora['mensaje']}")
    else:
        revisar(resultado["mensaje"] == "SIN_DETALLE" and ahora["mensaje"] == antes["mensaje"],
                "el mensaje no traia el detalle viejo tal cual: no se toco")
    evento = ultimo_evento()
    revisar(evento.get("TIPO_EVENTO") == co.TIPO_EVENTO and evento.get("ORIGEN") == "Prueba"
            and evento.get("ID_ENTIDAD") == codigo and "->" in str(evento.get("DETALLE"))
            and evento.get("ID_EVENTO") == resultado["evento"],
            f"evento {evento.get('ID_EVENTO')}: {evento.get('DETALLE')}")
    revisar(precios_vendidos(nombre) == precios_antes,
            f"el precio de las ventas ya hechas no cambio (Regla #40): {len(precios_antes[0])} de despacho y "
            f"{len(precios_antes[1])} digitales")

    paso("6) Lo mismo otra vez")
    huella_antes = huella(COPIA_TEMPORAL)
    resultado = guardar()
    revisar(resultado["estado"] == "SIN_CAMBIO" and huella(COPIA_TEMPORAL) == huella_antes,
            "la misma composicion: SIN_CAMBIO, no se escribe nada")

    paso("7) Venta DESPUES del cambio de cantidad")
    if puede_vender:
        venta = vender(codigo, 1)
        sumar(p["producto"], cantidad_nueva)
        sumar(r["producto"], r["cantidad"])
        revisar(round(venta["precio"], 2) == precio_esperado, f"1 x {nombre} se vendio al precio nuevo ({venta['precio']:,.2f})")
        inventario_como_se_espera(f"la venta nueva rebajo {cantidad_nueva:g} {p['producto']} (la cantidad nueva); lo anterior no se movio")
        revisar(precios_vendidos(nombre)[0][:-1] == precios_antes[0], "las ventas anteriores conservan su precio")
    else:
        print("       (sin ventas nuevas en esta copia)")

    paso("8) Cambia el PRODUCTO de reventa")
    antes = composicion(codigo)
    p, r = antes["componentes"]
    rango_antes, filas_antes = tabla_composicion()
    resultado = guardar(producto_2=otro["codigo"], cantidad_2=2)
    ahora = composicion(codigo)
    p2, r2 = ahora["componentes"]
    rango_ahora, filas_ahora = tabla_composicion()
    revisar(r2["producto"] == otro["nombre"] and r2["cantidad"] == 2 and r2["fila"] == r["fila"]
            and r2["congelados"] == r["vendidos"] and r2["rebajadas"] == 0 and r2["unidades"] == 0,
            f"la fila {r2['fila']} pasa de {r['producto']} a 2 x {otro['nombre']}: empieza en 0 "
            f"(G = {r2['congelados']:g} combos congelados, H = 0)")
    if r["unidades"]:
        historia = [h for h in ahora["historia"] if h["producto"] == r["producto"]]
        revisar(filas_ahora == filas_antes + 1 and len(historia) == 1 and historia[0]["cantidad"] == 0
                and historia[0]["unidades"] == r["unidades"],
                f"fila de historia {historia[0]['fila'] if historia else '?'}: {r['producto']}, cantidad 0, "
                f"{r['unidades']:g} unidades ya rebajadas; la tabla pasa de {rango_antes} a {rango_ahora}")
    else:
        revisar(filas_ahora == filas_antes, f"{r['producto']} no tenia unidades rebajadas: no hizo falta fila de historia")
    revisar((p2["cantidad"], p2["congelados"], p2["rebajadas"], p2["unidades"])
            == (p["cantidad"], p["congelados"], p["rebajadas"], p["unidades"]),
            f"la fila de {p['producto']} no se toco")
    inventario_como_se_espera("'Control de Inventario' quedo igual (lo ya rebajado sigue cargado al producto anterior)")
    revisar(ahora["detalle"] == co.texto_detalle(p2["cantidad"], p2["producto"], 2, otro["nombre"]),
            f"detalle: {ahora['detalle']} | precio {ahora['precio']:,.2f}")

    paso("9) Venta DESPUES del cambio de producto")
    if puede_vender:
        vender(codigo, 1)
        sumar(p2["producto"], p2["cantidad"])
        sumar(otro["nombre"], 2)
        inventario_como_se_espera(f"la venta rebajo 2 {otro['nombre']} y nada de {r['producto']}")
    else:
        print("       (sin ventas nuevas en esta copia)")

    paso("10) Volver al producto anterior y cambiar otra vez: la historia se acumula")
    guardar(producto_2=r["producto"], cantidad_2=r["cantidad"])
    if puede_vender:
        vender(codigo, 1)
        sumar(p2["producto"], p2["cantidad"])
        sumar(r["producto"], r["cantidad"])
    _, filas_medio = tabla_composicion()
    historia_antes = {h["producto"]: h["unidades"] for h in composicion(codigo)["historia"]}
    vigente = composicion(codigo)["componentes"][1]
    guardar(producto_2=otro["codigo"], cantidad_2=2)
    _, filas_fin = tabla_composicion()
    historia_ahora = {h["producto"]: h["unidades"] for h in composicion(codigo)["historia"]}
    if r["producto"] in historia_antes:
        revisar(filas_fin == filas_medio
                and historia_ahora[r["producto"]] == historia_antes[r["producto"]] + vigente["unidades"],
                f"{r['producto']} ya tenia fila de historia: se acumulo ahi ({historia_antes[r['producto']]:g} + "
                f"{vigente['unidades']:g} = {historia_ahora[r['producto']]:g}) y la tabla no crecio")
    else:
        revisar(filas_fin == filas_medio + (1 if vigente["unidades"] else 0),
                "sin historia previa de ese producto: se crea la fila solo si habia unidades rebajadas")
    inventario_como_se_espera("'Control de Inventario' sigue cuadrando")

    paso("11) Devolver todo")
    resultado = guardar(cantidad_1=inicio["componentes"][0]["cantidad"], producto_2=inicio["componentes"][1]["producto"],
                        cantidad_2=inicio["componentes"][1]["cantidad"])
    final = composicion(codigo)
    revisar([(c["producto"], c["cantidad"]) for c in final["componentes"]]
            == [(c["producto"], c["cantidad"]) for c in inicio["componentes"]]
            and round(final["precio"], 2) == round(inicio["precio"], 2),
            f"{nombre} volvio a su composicion ({final['detalle']}) y a su precio ({final['precio']:,.2f})")
    inventario_como_se_espera("el inventario solo se movio por las ventas de la prueba, cada una con su composicion")
    if not elegido["activo"] and puede_vender:
        cc.actualizar_combo(codigo, origen="Prueba", **{campo: elegido[campo] for campo in cc.CAMPOS})

    paso("12) Cierre")
    hoja = leer_excel(solo_valores=False)[co.HOJA]
    revisar(all(cc._es_formula(hoja[c["celdas"][clave]].value) for c in combos_inicio for clave in ("estado", "precio")),
            "el estado (A) y el precio (E) de todos los combos siguen siendo formulas")
    revisar([c["codigo"] for c in cc.leer_combos()] == [c["codigo"] for c in combos_inicio],
            f"los combos siguen siendo {len(combos_inicio)}: las filas de historia no se confunden con combos")
    propio_inicio = inicio["componentes"][0]
    revisar(Catalogo(leer_excel()).resolver(propio_inicio["producto"])["codigo"] == propio_inicio["codigo_producto"],
            f"el catalogo sigue encontrando '{propio_inicio['producto']}' por nombre como producto ({propio_inicio['codigo_producto']})")
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
