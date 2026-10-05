"""Microservicio: COMPOSICION de los combos (Grupo 3, Decisiones Pendientes #85, opcion 3).

Que lleva cada combo. Diseno confirmado por el usuario (5-oct-2026):

  * Un combo tiene DOS componentes: el primero de produccion propia (Tipo 'Propio' en la TABLA
    CRUCE MAESTRA: Baguette o CupCake) y el segundo de reventa (Tipo 'Tercero').
  * Cantidad de cada uno: numero entero de 1 a 10.
  * Se puede cambiar AUNQUE el combo ya tenga ventas. Lo ya vendido no se modifica (Regla #40):
    el dinero ya esta protegido (cada venta guarda su precio como numero) y el inventario se
    protege CONGELANDO lo ya rebajado en las columnas G y H que creo la migracion m014.

La composicion vive en TRES lugares de 'Tabla Combos Dinamicos' que deben coincidir, y esta
funcion los escribe en UN solo guardado:

  1) TablaComposicionCombos (A15:H..), dos filas vigentes por combo (las de cantidad mayor que 0).
     'Control de Inventario' columna I suma su columna F por producto.
       - Cambia solo la CANTIDAD:  G := combos vendidos hasta hoy,  H := unidades ya rebajadas,
         y despues D := cantidad nueva.  F = H + D x (E - G) no se mueve; la siguiente venta suma
         la cantidad nueva.
       - Cambia el PRODUCTO: lo ya rebajado se queda cargado al producto anterior en una FILA DE
         HISTORIA (producto anterior, cantidad 0, H = lo ya rebajado; si ya existe una para ese combo
         y producto, se acumula en ella). La fila vigente pasa al producto nuevo: C enlace a la TCM,
         D cantidad nueva, G := combos vendidos, H := 0. Si no habia nada rebajado no hace falta fila
         de historia.
  2) La formula del Precio de Oferta (columna E del combo), con las cantidades y las filas nuevas
     de la TCM. El descuento (columna L) sigue aplicando solo al componente propio (Regla #28).
  3) El texto del Detalle (columna D), con los nombres de la TCM (Regla #26):
     '4 Baguette + 1 Gaseosa Coca Cola Regular - 600 ml'.
  Ademas: si el Mensaje SaaS / Bot (columna K) trae el detalle viejo tal cual, se cambia por el nuevo.

Antes de escribir se comprueba que los tres lugares COINCIDEN hoy (la formula del precio usa las
mismas cantidades y productos que la tabla). Si no coinciden, se detiene y no escribe nada.

LIMITE CONOCIDO (#92, decision del usuario): se congela por CONTADOR de combos vendidos, no por
fecha. Si despues de un cambio se anula una venta hecha ANTES del cambio, el inventario devuelve
las unidades de la composicion nueva y no las de la vieja.

Reglas del proyecto que se respetan:
  * El combo se ubica por su CODIGO (co01), las columnas por su ENCABEZADO y los productos por su
    nombre o codigo de la TCM. Nada por numero fijo de fila.
  * Si nada cambia, no se sube nada (estado SIN_CAMBIO).
  * Un guardado = UNA escritura y UN evento, con lo anterior -> lo nuevo y lo congelado.
  * El inventario lo sigue rebajando el Excel con sus formulas (Regla #27).
  * No cambia el detonador ni el descuento: por eso aqui no se revisa 'Activacion de Promocion'.

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo3_integraciones.combos_composicion --local --ver
"""
import os
import re
import sys

# --local: igual que combos_config.py y las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import editar_excel, leer_excel                                   # noqa: E402
from conexion_drive.tablas import (agregar_fila, columnas_de_tabla, filas_de_datos,   # noqa: E402
                                   obtener_tabla, registrar_evento)
from . import combos_config as cc                                                     # noqa: E402

HOJA = cc.HOJA
TABLA = "TablaComposicionCombos"
HOJA_TCM, TABLA_TCM = "TABLA CRUCE MAESTRA", "TablaProductos"
TIPO_EVENTO = "COMPOSICION_COMBO"
PROPIO, TERCERO = "Propio", "Tercero"
ROLES = ((PROPIO, "Producción propia"), (TERCERO, "Reventa"))
CANTIDAD_MINIMA, CANTIDAD_MAXIMA = 1, 10

# Encabezados de TablaComposicionCombos (los dos ultimos los creo la migracion m014).
COL = {"codigo": "Codigo combo", "combo": "Nombre del combo", "producto": "Producto",
       "cantidad": "Cantidad por combo", "vendidos": "Combos vendidos", "unidades": "Unidades del producto",
       "congelados": "Combos ya congelados", "rebajadas": "Unidades ya rebajadas"}
# Encabezados de TablaProductos (TABLA CRUCE MAESTRA).
TCM_NOMBRE, TCM_CODIGO, TCM_PRECIO, TCM_TIPO = "Nombre tecnico completo", "Codigo", "Precio", "Tipo"

ValorInvalido = cc.ValorInvalido
ComboDesconocido = cc.ComboDesconocido
_texto, _numero, _es_formula = cc._texto, cc._numero, cc._es_formula


# ------------------------------------------------------------------ validaciones
def normalizar_cantidad(valor, nombre="Cantidad"):
    """Lo digitado como numero entero de 1 a 10. Si no, ValorInvalido."""
    numero = cc._a_numero(valor, nombre, "3")
    if numero != int(numero) or not CANTIDAD_MINIMA <= numero <= CANTIDAD_MAXIMA:
        raise ValorInvalido(f"{nombre}: debe ser un numero entero de {CANTIDAD_MINIMA} a {CANTIDAD_MAXIMA} "
                            f"(llego {numero:g}). No se escribio nada.")
    return int(numero)


def texto_detalle(cantidad_1, nombre_1, cantidad_2, nombre_2):
    """El texto de la columna D: '3 Baguette + 1 Leche Semidescremada Dos Pinos -1 L'."""
    return f"{cantidad_1:g} {_texto(nombre_1)} + {cantidad_2:g} {_texto(nombre_2)}"


# ------------------------------------------------------------------ lectura
def _productos(wb_valores):
    """Los productos de la TCM: ([{'nombre', 'codigo', 'tipo', 'precio', 'fila'}], letra nombre, letra precio)."""
    ws, tabla = obtener_tabla(wb_valores, HOJA_TCM, TABLA_TCM)
    columnas = columnas_de_tabla(tabla)
    faltan = [c for c in (TCM_NOMBRE, TCM_CODIGO, TCM_PRECIO, TCM_TIPO) if c not in columnas]
    if faltan:
        raise ValorInvalido(f"En {TABLA_TCM} faltan las columnas {faltan}. No se escribio nada.")
    productos = []
    for fila in filas_de_datos(ws, tabla):
        nombre = _texto(ws[f"{columnas[TCM_NOMBRE]}{fila}"].value)
        if nombre:
            productos.append({"nombre": nombre, "codigo": _texto(ws[f"{columnas[TCM_CODIGO]}{fila}"].value),
                              "tipo": _texto(ws[f"{columnas[TCM_TIPO]}{fila}"].value),
                              "precio": _numero(ws[f"{columnas[TCM_PRECIO]}{fila}"].value), "fila": fila})
    return productos, columnas[TCM_NOMBRE], columnas[TCM_PRECIO]


def _columnas_composicion(tabla):
    columnas = columnas_de_tabla(tabla)
    faltan = [texto for texto in COL.values() if texto not in columnas]
    if faltan:
        raise ValorInvalido(f"En {TABLA} faltan las columnas {faltan}: hay que aplicar la migracion m014. "
                            "No se escribio nada.")
    return {clave: columnas[texto] for clave, texto in COL.items()}


def _filas_composicion(wb_valores):
    """Todas las filas de TablaComposicionCombos como VALORES: [{'fila', 'codigo', 'combo', 'producto', ...}]."""
    ws, tabla = obtener_tabla(wb_valores, HOJA, TABLA)
    letras = _columnas_composicion(tabla)
    filas = []
    for fila in filas_de_datos(ws, tabla):
        leido = {clave: ws[f"{letra}{fila}"].value for clave, letra in letras.items()}
        if _texto(leido["codigo"]) or _texto(leido["producto"]):
            filas.append({"fila": fila, "codigo": _texto(leido["codigo"]), "combo": _texto(leido["combo"]),
                          "producto": _texto(leido["producto"]),
                          **{c: leido[c] for c in ("cantidad", "vendidos", "unidades", "congelados", "rebajadas")},
                          "celdas": {clave: f"{letra}{fila}" for clave, letra in letras.items()}})
    return filas


def _armar(filas_combo, productos):
    """Separa las filas de UN combo en sus dos componentes vigentes y su historia. Si no tiene la forma
    esperada devuelve (None, historia, motivo)."""
    por_nombre = {p["nombre"]: p for p in productos}
    vigentes = [f for f in filas_combo if (_numero(f["cantidad"]) or 0) > 0]
    historia = [f for f in filas_combo if not (_numero(f["cantidad"]) or 0) > 0]
    desconocidos = [f["producto"] for f in filas_combo if f["producto"] not in por_nombre]
    if desconocidos:
        return None, historia, f"el producto {desconocidos[0]!r} de la tabla de composicion no esta en la {HOJA_TCM}"
    componentes = []
    for tipo, _ in ROLES:
        del_tipo = [f for f in vigentes if por_nombre[f["producto"]]["tipo"] == tipo]
        if len(del_tipo) != 1:
            return None, historia, (f"se esperaba UN componente vigente de tipo '{tipo}' y hay {len(del_tipo)} "
                                    f"(filas {[f['fila'] for f in vigentes]})")
        componentes.append(dict(del_tipo[0], tipo=tipo, codigo_producto=por_nombre[del_tipo[0]["producto"]]["codigo"],
                                fila_tcm=por_nombre[del_tipo[0]["producto"]]["fila"]))
    if len(vigentes) != 2:
        return None, historia, f"se esperaban 2 filas vigentes y hay {len(vigentes)}"
    numeros = [c[k] for c in componentes for k in ("cantidad", "vendidos", "unidades", "congelados", "rebajadas")]
    if any(_numero(n) is None for n in numeros):
        return None, historia, "la tabla de composicion tiene celdas sin calcular (el Excel debe recalcularse)"
    return componentes, historia, ""


def leer_composicion(wb_valores=None):
    """La composicion de cada combo, tal como esta en el Excel ya calculado (solo lectura).

    Devuelve {'combos': [{'codigo', 'nombre', 'fila', 'detalle', 'precio', 'ajustable', 'motivo',
                          'componentes': [propio, reventa], 'historia': [...]}],
              'propios': [productos Tipo 'Propio'], 'terceros': [productos Tipo 'Tercero']}
    Cada componente: {'tipo', 'producto', 'codigo_producto', 'cantidad', 'vendidos', 'unidades',
                      'congelados', 'rebajadas', 'fila', 'celdas': {'producto': 'C16', 'cantidad': 'D16', ...}}
    """
    wb_valores = wb_valores or leer_excel()
    productos, _, _ = _productos(wb_valores)
    filas = _filas_composicion(wb_valores)
    combos = []
    for combo in cc.leer_combos(wb_valores):
        componentes, historia, motivo = _armar([f for f in filas if f["codigo"].lower() == combo["codigo"].lower()],
                                               productos)
        combos.append({"codigo": combo["codigo"], "nombre": combo["nombre"], "fila": combo["fila"],
                       "detalle": combo["detalle"], "precio": combo["precio"], "mensaje": combo["mensaje"],
                       "ajustable": componentes is not None, "motivo": motivo,
                       "componentes": componentes or [], "historia": historia,
                       "celdas": {"detalle": combo["celdas"]["detalle"], "precio": combo["celdas"]["precio"]}})
    return {"combos": combos, "propios": [p for p in productos if p["tipo"] == PROPIO],
            "terceros": [p for p in productos if p["tipo"] == TERCERO]}


# ------------------------------------------------------------------ formula del precio
def formula_precio(fila_combo, letra_descuento, letra_precio, cantidad_1, fila_tcm_1, cantidad_2, fila_tcm_2):
    """La formula del Precio de Oferta: (propio) con el descuento de la fila + (reventa) a su precio."""
    d = f"${letra_descuento}{fila_combo}"
    return (f"=({cantidad_1}*'{HOJA_TCM}'!{letra_precio}{fila_tcm_1})*(1-IF({d}>1,{d}/100,{d}))"
            f"+{cantidad_2}*'{HOJA_TCM}'!{letra_precio}{fila_tcm_2}")


def leer_formula_precio(formula, fila_combo, letra_descuento, letra_precio):
    """De la formula actual saca (cantidad 1, fila TCM 1, cantidad 2, fila TCM 2). None si no tiene esa forma."""
    limpia = re.sub(r"\s+", "", str(formula or ""))          # sin espacios: tambien dentro del nombre de la hoja
    hoja = re.escape(re.sub(r"\s+", "", f"'{HOJA_TCM}'!"))
    d = re.escape(f"${letra_descuento}{fila_combo}")
    patron = (rf"^=\((\d+)\*{hoja}\$?{letra_precio}\$?(\d+)\)\*\(1-IF\({d}>1,{d}/100,{d}\)\)"
              rf"\+\(?(?:(\d+)\*)?{hoja}\$?{letra_precio}\$?(\d+)\)?$")
    encontrado = re.match(patron, limpia)
    if not encontrado:
        return None
    n1, f1, n2, f2 = encontrado.groups()
    return int(n1), int(f1), int(n2 or 1), int(f2)


# ------------------------------------------------------------------ escritura
def actualizar_composicion(codigo, producto_1, cantidad_1, producto_2, cantidad_2, origen="SaaS"):
    """Guarda la composicion de UN combo: componente 1 (produccion propia) y componente 2 (reventa).

    codigo: 'co01' (sin distinguir mayusculas). Los productos llegan por CODIGO ('P01') o por su nombre
    exacto de la TABLA CRUCE MAESTRA; las cantidades, como las digito la persona.
    Devuelve {'hubo_cambios', 'estado', 'codigo', 'nombre', 'fila', 'antes', 'despues' (textos de la
              composicion), 'cambios': [{'titulo', 'celda', 'antes', 'despues'}],
              'congelado': [{'producto', 'combos', 'unidades', 'fila', 'fila_historia'}],
              'mensaje': 'ACTUALIZADO' / 'SIN_DETALLE' / 'MUY_LARGO' / 'IGUAL', 'evento'}.
    """
    buscado = _texto(codigo).lower()

    def _editar(wb, wb_valores):
        ws, ws_valores = wb[HOJA], wb_valores[HOJA]
        fila_encabezado, columnas = cc._columnas(ws)
        filas = [f for f in cc._filas_de_combos(ws, fila_encabezado, columnas)
                 if _texto(ws.cell(f, columnas["codigo"]).value).lower() == buscado]
        if len(filas) != 1:
            raise ComboDesconocido(f"El combo '{codigo}' aparece {len(filas)} veces en '{HOJA}' "
                                   "(debe aparecer una sola vez). No se escribio nada.")
        fila_combo = filas[0]
        celda = {clave: ws.cell(fila_combo, columna) for clave, columna in columnas.items()}
        nombre_combo, codigo_combo = _texto(celda["nombre"].value), _texto(celda["codigo"].value)

        # Lo pedido: dos productos de la TCM (uno propio y uno de reventa) y dos cantidades enteras.
        productos, letra_nombre, letra_precio = _productos(wb_valores)

        def producto_pedido(valor, tipo, titulo):
            pedido = _texto(valor)
            hallados = [p for p in productos if pedido and (p["codigo"].lower() == pedido.lower() or p["nombre"] == pedido)]
            if len(hallados) != 1:
                raise ValorInvalido(f"{titulo}: {valor!r} no es un producto de la {HOJA_TCM}. No se escribio nada.")
            if hallados[0]["tipo"] != tipo:
                opciones = ", ".join(p["nombre"] for p in productos if p["tipo"] == tipo)
                raise ValorInvalido(f"{titulo}: '{hallados[0]['nombre']}' es de tipo '{hallados[0]['tipo']}' y aqui va "
                                    f"un producto de tipo '{tipo}' ({opciones}). No se escribio nada.")
            return hallados[0]

        nuevos = [(producto_pedido(producto_1, PROPIO, "Componente 1 (producción propia)"),
                   normalizar_cantidad(cantidad_1, "Cantidad del componente 1")),
                  (producto_pedido(producto_2, TERCERO, "Componente 2 (reventa)"),
                   normalizar_cantidad(cantidad_2, "Cantidad del componente 2"))]

        # Lo que hay hoy: la tabla de composicion (valores ya calculados) ...
        tabla = ws.tables.get(TABLA)
        if tabla is None:
            raise ValorInvalido(f"No existe {TABLA} en '{HOJA}'. No se escribio nada.")
        letras = _columnas_composicion(tabla)
        filas_combo = [f for f in _filas_composicion(wb_valores) if f["codigo"].lower() == buscado]
        componentes, historia, motivo = _armar(filas_combo, productos)
        if componentes is None:
            raise ValorInvalido(f"La composicion de {nombre_combo} ({codigo_combo}) no tiene la forma esperada: "
                                f"{motivo}. No se escribio nada.")
        # ... y la formula del precio, que debe decir lo mismo que la tabla.
        if not _es_formula(celda["precio"].value):
            raise ValorInvalido(f"'{HOJA}'!{celda['precio'].coordinate} deberia ser la formula del precio. No se escribio nada.")
        letra_descuento = celda["descuento"].column_letter
        en_formula = leer_formula_precio(celda["precio"].value, fila_combo, letra_descuento, letra_precio)
        en_tabla = (componentes[0]["cantidad"], componentes[0]["fila_tcm"], componentes[1]["cantidad"], componentes[1]["fila_tcm"])
        if en_formula is None or en_formula != tuple(int(n) for n in en_tabla):
            raise ValorInvalido(
                f"La formula del precio '{HOJA}'!{celda['precio'].coordinate} ({celda['precio'].value}) y la tabla de "
                f"composicion (filas {componentes[0]['fila']} y {componentes[1]['fila']}) no dicen lo mismo. "
                "Hay que revisarlo antes de cambiar la composicion. No se escribio nada.")
        for clave in ("detalle", "mensaje"):
            if _es_formula(celda[clave].value):
                raise ValorInvalido(f"'{HOJA}'!{celda[clave].coordinate} es una formula: el robot no la sobrescribe.")

        antes = texto_detalle(componentes[0]["cantidad"], componentes[0]["producto"],
                              componentes[1]["cantidad"], componentes[1]["producto"])
        despues = texto_detalle(nuevos[0][1], nuevos[0][0]["nombre"], nuevos[1][1], nuevos[1][0]["nombre"])
        base = {"codigo": codigo_combo, "nombre": nombre_combo, "fila": fila_combo, "antes": antes, "despues": despues}
        if antes == despues:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", "cambios": [], "congelado": [], **base}

        cambios, congelado = [], []

        def escribir(coordenada, valor, titulo, hoja=ws):
            anterior = hoja[coordenada].value
            hoja[coordenada].value = valor
            cambios.append({"titulo": titulo, "celda": coordenada, "antes": anterior, "despues": valor})

        # 1) Tabla de composicion: congelar lo ya rebajado y escribir lo nuevo.
        for actual, (producto, cantidad) in zip(componentes, nuevos):
            mismo_producto, misma_cantidad = actual["producto"] == producto["nombre"], actual["cantidad"] == cantidad
            if mismo_producto and misma_cantidad:
                continue
            c = actual["celdas"]
            for clave in ("cantidad", "congelados", "rebajadas"):
                if _es_formula(ws[c[clave]].value) or _numero(ws[c[clave]].value) is None:
                    raise ValorInvalido(f"'{HOJA}'!{c[clave]} deberia ser un numero y es {ws[c[clave]].value!r}. "
                                        "No se escribio nada.")
            vendidos, unidades = actual["vendidos"], actual["unidades"]
            dato = {"producto": actual["producto"], "combos": vendidos, "unidades": unidades, "fila": actual["fila"],
                    "fila_historia": None}
            if mismo_producto:
                # Solo la cantidad: lo vendido hasta hoy queda con las unidades viejas.
                escribir(c["congelados"], vendidos, f"{COL['congelados']} ({actual['producto']})")
                escribir(c["rebajadas"], unidades, f"{COL['rebajadas']} ({actual['producto']})")
                escribir(c["cantidad"], cantidad, f"{COL['cantidad']} ({actual['producto']})")
            else:
                # Otro producto: lo ya rebajado sigue cargado al producto anterior, en una fila de historia.
                if unidades:
                    previa = next((h for h in historia if h["producto"] == actual["producto"]), None)
                    if previa is not None:
                        acumulado = (_numero(ws[previa["celdas"]["rebajadas"]].value) or 0) + unidades
                        escribir(previa["celdas"]["rebajadas"], acumulado,
                                 f"{COL['rebajadas']} (historia de {actual['producto']})")
                        escribir(previa["celdas"]["congelados"], vendidos,
                                 f"{COL['congelados']} (historia de {actual['producto']})")
                        dato["fila_historia"] = previa["fila"]
                    else:
                        fila_nueva = agregar_fila(ws, tabla, {
                            COL["codigo"]: ws[c["codigo"]].value, COL["combo"]: ws[c["combo"]].value,
                            COL["producto"]: ws[c["producto"]].value, COL["cantidad"]: 0,
                            COL["congelados"]: vendidos, COL["rebajadas"]: unidades,
                        }, columnas_clave=[COL["producto"]])
                        # Las dos formulas de la fila, con su propio numero de fila.
                        formula_vendidos = re.sub(rf"(?<![$A-Za-z']){letras['combo']}{actual['fila']}(?!\d)",
                                                  f"{letras['combo']}{fila_nueva}", str(ws[c["vendidos"]].value))
                        ws[f"{letras['vendidos']}{fila_nueva}"] = formula_vendidos
                        ws[f"{letras['unidades']}{fila_nueva}"] = (
                            f"={letras['rebajadas']}{fila_nueva}+{letras['cantidad']}{fila_nueva}"
                            f"*({letras['vendidos']}{fila_nueva}-{letras['congelados']}{fila_nueva})")
                        cambios.append({"titulo": f"Fila de historia ({actual['producto']})",
                                        "celda": f"{letras['codigo']}{fila_nueva}:{letras['rebajadas']}{fila_nueva}",
                                        "antes": None, "despues": f"cantidad 0, {unidades:g} unidades ya rebajadas"})
                        dato["fila_historia"] = fila_nueva
                escribir(c["producto"], f"='{HOJA_TCM}'!${letra_nombre}${producto['fila']}",
                         f"{COL['producto']} (antes {actual['producto']})")
                cambios[-1]["despues"] = producto["nombre"]
                cambios[-1]["antes"] = actual["producto"]
                escribir(c["congelados"], vendidos, f"{COL['congelados']} ({producto['nombre']})")
                escribir(c["rebajadas"], 0, f"{COL['rebajadas']} ({producto['nombre']})")
                if not misma_cantidad:
                    escribir(c["cantidad"], cantidad, f"{COL['cantidad']} ({producto['nombre']})")
            congelado.append(dato)

        # 2) La formula del precio, 3) el detalle y, si trae el detalle viejo tal cual, el mensaje.
        escribir(celda["precio"].coordinate,
                 formula_precio(fila_combo, letra_descuento, letra_precio, nuevos[0][1], nuevos[0][0]["fila"],
                                nuevos[1][1], nuevos[1][0]["fila"]), cc.ENCABEZADOS["precio"])
        detalle_viejo = _texto(celda["detalle"].value)
        escribir(celda["detalle"].coordinate, despues, cc.ENCABEZADOS["detalle"])
        mensaje_viejo = "" if celda["mensaje"].value is None else str(celda["mensaje"].value)
        if not detalle_viejo or detalle_viejo not in mensaje_viejo:
            estado_mensaje = "SIN_DETALLE"
        else:
            mensaje_nuevo = mensaje_viejo.replace(detalle_viejo, despues)
            if mensaje_nuevo == mensaje_viejo:
                estado_mensaje = "IGUAL"
            elif len(mensaje_nuevo) > cc.LARGO_MAXIMO_MENSAJE:
                estado_mensaje = "MUY_LARGO"
            else:
                estado_mensaje = "ACTUALIZADO"
                escribir(celda["mensaje"].coordinate, mensaje_nuevo, cc.ENCABEZADOS["mensaje"])

        partes = [f"{nombre_combo} ({codigo_combo}), fila {fila_combo}: '{antes}' -> '{despues}'"]
        for dato in congelado:
            texto = (f"congelado {dato['producto']} (fila {dato['fila']}): {dato['combos']:g} combos, "
                     f"{dato['unidades']:g} unidades")
            if dato["fila_historia"]:
                texto += f", en la fila de historia {dato['fila_historia']}"
            partes.append(texto)
        partes.append(f"precio {celda['precio'].coordinate}, detalle {celda['detalle'].coordinate}"
                      + {"ACTUALIZADO": f" y mensaje {celda['mensaje'].coordinate} reescritos",
                         "MUY_LARGO": f" reescritos; mensaje {celda['mensaje'].coordinate} NO (pasaria de "
                                      f"{cc.LARGO_MAXIMO_MENSAJE} caracteres)",
                         "SIN_DETALLE": f" reescritos; mensaje {celda['mensaje'].coordinate} sin tocar (no traia el detalle)",
                         "IGUAL": " reescritos"}[estado_mensaje])
        partes.append("Regla #40: lo ya vendido no cambia")
        evento = registrar_evento(
            wb, modulo="Combos", tipo_evento=TIPO_EVENTO, entidad=HOJA, id_entidad=codigo_combo,
            accion="UPDATE", estado="OK", origen=origen, valor_referencia=len(cambios), detalle=" | ".join(partes),
        )
        return {"hubo_cambios": True, "estado": "OK", "cambios": cambios, "congelado": congelado,
                "mensaje": estado_mensaje, "evento": evento, **base}

    return editar_excel(_editar, descripcion=f"Composicion del combo {codigo}")


if __name__ == "__main__":
    argumentos = [a for a in sys.argv[1:] if a != "--local"]
    if argumentos != ["--ver"]:
        sys.exit("Uso: py -m microservicios.grupo3_integraciones.combos_composicion [--local] --ver")
    leido = leer_composicion()
    for combo in leido["combos"]:
        print(f"  fila {combo['fila']}  {combo['codigo']}  {combo['nombre']:<8}  precio {combo['precio']:,.2f}  | {combo['detalle']}")
        if not combo["ajustable"]:
            print(f"      NO AJUSTABLE: {combo['motivo']}")
        for comp in combo["componentes"]:
            print(f"      fila {comp['fila']}  {comp['tipo']:<8} {comp['cantidad']:g} x {comp['producto']}  "
                  f"(vendidos {comp['vendidos']:g}, unidades {comp['unidades']:g}, congelados {comp['congelados']:g}, "
                  f"ya rebajadas {comp['rebajadas']:g})")
        for fila in combo["historia"]:
            print(f"      fila {fila['fila']}  historia  {fila['producto']}: {(_numero(fila['unidades']) or 0):g} unidades ya rebajadas")
    print("  Propios: " + ", ".join(p["nombre"] for p in leido["propios"]))
    print("  Reventa: " + ", ".join(p["nombre"] for p in leido["terceros"]))
