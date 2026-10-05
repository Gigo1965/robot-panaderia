"""Microservicio: CONFIGURACION de los combos (Grupo 3, Decisiones Pendientes #65).

Mapa de celdas 'Tabla combos dinamicos': en cada combo (filas 5 a 8 de 'Tabla Combos Dinamicos')
la administracion configura seis celdas; lo demas es formula o dato fijo:

    G  Tipo detonador    lista TEMPERATURA / LLUVIA
    H  Operador          lista <, <=, >, >=          (solo con TEMPERATURA)
    I  Valor limite      TEMPERATURA: de 0 a 50 grados, con un decimal como maximo
                         LLUVIA: Lluvia / Despejado (lo mismo que escribe el robot en 'NODO INTEGRAL'!C5)
    J  Prioridad         entero de 1 a 9 (dato de referencia: hoy ninguna formula ni el codigo la leen)
    K  Mensaje SaaS/Bot  texto de 1 a 150 caracteres
    L  Descuento %       de 0 a 10, con un decimal como maximo (Regla #28: lo digita la administracion)

Solo lectura aqui: A (Detonador, formula), C (Nombre, es la LLAVE), D (Detalle), E (Precio de
Oferta, formula) y F (codigo). La columna B (texto de referencia del criterio) no la digita nadie:
la REESCRIBE el robot para que coincida con H e I cada vez que guarda un cambio.

Decisiones del usuario (4 y 5-oct-2026):
  * Detonadores permitidos: TEMPERATURA y LLUVIA. FERIADO y DIA_PAGO NO: esos dias hay descuento
    general y no hay combos (Regla #8).
  * Tope del descuento: 10 %.
  * Un formulario por combo, solo administracion; un guardado = UNA escritura y UN evento.
  * Despues de guardar, el robot revisa 'Activacion de Promocion' ('NODO INTEGRAL'!C16, #84).

Detalle de construccion (por la formula del precio, columna E): el descuento se lee con
IF(L>1, L/100, L). Un numero MAYOR que 1 es un porcentaje (10 = 10 %); uno de 1 o MENOS es una
fraccion (0.1 = 10 %, y un 1 seria 100 %). Por eso el robot guarda 2, 5 o 10 tal cual, y los
porcentajes de 1 o menos como fraccion (1 % -> 0.01; 0.5 % -> 0.005). Asi nunca queda un 100 %.

Reglas del proyecto que se respetan:
  * Es CONFIGURACION: se actualizan solo las celdas que cambiaron y queda un evento en
    RegistroEventos con el VALOR ANTERIOR -> VALOR NUEVO de cada una.
  * Nunca se escribe sobre una formula. El combo se ubica por su CODIGO (co01) y las columnas por
    su ENCABEZADO (como el catalogo), no por numero fijo.
  * Si nada cambia en G a L, no se sube nada (estado SIN_CAMBIO).
  * Lo ya vendido no cambia (Regla #40): el precio de cada venta esta guardado como numero.
  * La COMPOSICION del combo (que productos lleva) no se cambia aqui: va en otra funcion (#85).

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo3_integraciones.combos_config --local --ver
"""
import math
import os
import re
import sys

# --local: igual que nodo_manual.py y las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import editar_excel, leer_excel                                   # noqa: E402
from conexion_drive.tablas import registrar_evento                                    # noqa: E402
from . import activacion_promocion as ap                                              # noqa: E402

HOJA = "Tabla Combos Dinamicos"
TIPO_EVENTO = "CONFIG_COMBO"
TEMPERATURA, LLUVIA = "TEMPERATURA", "LLUVIA"
TIPOS = (TEMPERATURA, LLUVIA)
OPERADORES = ("<", "<=", ">", ">=")
CLIMAS = ("Lluvia", "Despejado")                 # los dos textos que el robot escribe en 'NODO INTEGRAL'!C5
VALOR_MINIMO, VALOR_MAXIMO = 0, 50               # grados
PRIORIDAD_MINIMA, PRIORIDAD_MAXIMA = 1, 9
LARGO_MAXIMO_MENSAJE = 150
DESCUENTO_MAXIMO = 10                            # por ciento

# Encabezados de la hoja (fila 4 hoy). clave -> texto del encabezado, sin espacios a los lados.
ENCABEZADOS = {
    "estado": "Detonador", "clima": "Clima", "nombre": "Nombre del Combo (Key)",
    "detalle": "Detalle del Combo (Descripción para el Robot)", "precio": "Precio de Oferta (Fórmula)",
    "codigo": "codigo", "tipo": "Tipo detonador", "operador": "Operador", "valor": "Valor límite",
    "prioridad": "Prioridad", "mensaje": "Mensaje SaaS / Bot", "descuento": "Descuento %",
}
# Las seis celdas que configura la administracion, en el orden de la hoja (G a L).
CAMPOS = ("tipo", "operador", "valor", "prioridad", "mensaje", "descuento")


class ValorInvalido(Exception):
    """Dato rechazado por una validacion: no se escribio nada."""


class ComboDesconocido(ValorInvalido):
    """El codigo pedido no es un combo de la hoja."""


def _texto(valor):
    return "" if valor is None or isinstance(valor, bool) else str(valor).strip()


def _numero(valor):
    return valor if isinstance(valor, (int, float)) and not isinstance(valor, bool) else None


def _es_formula(valor):
    return (isinstance(valor, str) and valor.startswith("=")) or type(valor).__name__ == "ArrayFormula"


def _a_numero(valor, nombre, ejemplo):
    """Convierte lo digitado en numero. Acepta coma o punto decimal; rechaza lo demas."""
    if isinstance(valor, bool):
        raise ValorInvalido(f"{nombre}: {valor!r} no es un numero.")
    if _numero(valor) is not None:
        numero = float(valor)
    else:
        texto = _texto(valor).replace("%", "").replace("°", "").replace(" ", "")
        if "," in texto and "." in texto:
            raise ValorInvalido(f"{nombre}: '{valor}' no se entiende. Escriba un numero, por ejemplo {ejemplo}.")
        try:
            numero = float(texto.replace(",", "."))
        except ValueError:
            raise ValorInvalido(f"{nombre}: {valor!r} no es un numero. Escriba, por ejemplo, {ejemplo}.") from None
    if math.isnan(numero) or math.isinf(numero):
        raise ValorInvalido(f"{nombre}: {valor!r} no es un numero. Escriba, por ejemplo, {ejemplo}.")
    return numero


def _entero_si_se_puede(numero):
    return int(numero) if numero == int(numero) else numero


def _un_decimal(numero, nombre):
    if round(numero, 1) != round(numero, 6):
        raise ValorInvalido(f"{nombre}: se acepta un decimal como maximo (llego {numero:g}). No se escribio nada.")
    return _entero_si_se_puede(round(numero, 1))


def descuento_a_celda(porcentaje):
    """Como se guarda un porcentaje en la columna L (ver el detalle de construccion de arriba)."""
    return porcentaje if porcentaje > 1 or porcentaje == 0 else round(porcentaje / 100, 6)


def descuento_de_celda(valor):
    """El porcentaje que la formula del precio entiende en la columna L. Vacia = 0 (sin descuento)."""
    numero = _numero(valor)
    if numero is None:
        return 0
    return _entero_si_se_puede(round(numero if numero > 1 else numero * 100, 4))


def texto_criterio(tipo, operador, valor):
    """El texto de referencia de la columna B: '<= 30' o 'Lluvia'."""
    return f"{operador} {valor:g}" if tipo == TEMPERATURA else str(valor)


def normalizar_configuracion(tipo, operador, valor, prioridad, mensaje, descuento, operador_actual=None):
    """Valida lo digitado y lo devuelve como se guarda en las celdas G a L. Si no, ValorInvalido.

    Con LLUVIA el operador no se usa: se conserva el que ya tiene la celda (operador_actual).
    """
    tipo_n = _texto(tipo).upper()
    if tipo_n in ("FERIADO", "DIA_PAGO"):
        raise ValorInvalido(
            f"Tipo detonador: {tipo_n} no se permite. En feriado y en dia de pago hay descuento general "
            "y no hay combos (Regla #8). No se escribio nada.")
    if tipo_n not in TIPOS:
        raise ValorInvalido(f"Tipo detonador: {tipo!r} no es una opcion. Las opciones son "
                            f"{' o '.join(TIPOS)}. No se escribio nada.")

    if tipo_n == TEMPERATURA:
        operador_n = re.sub(r"\s+", "", _texto(operador))
        if operador_n not in OPERADORES:
            raise ValorInvalido(f"Operador: {operador!r} no es una opcion. Las opciones son "
                                f"{', '.join(OPERADORES)}. No se escribio nada.")
        numero = _a_numero(valor, "Valor limite", "30")
        if not VALOR_MINIMO <= numero <= VALOR_MAXIMO:
            raise ValorInvalido(f"Valor limite: la temperatura debe estar entre {VALOR_MINIMO} y {VALOR_MAXIMO} "
                                f"grados (llego {numero:g}). No se escribio nada.")
        valor_n = _un_decimal(numero, "Valor limite")
    else:
        operador_n = operador_actual                      # no se usa con LLUVIA: la celda queda como esta
        valor_n = next((c for c in CLIMAS if c.lower() == _texto(valor).lower()), None)
        if valor_n is None:
            raise ValorInvalido(f"Valor limite: con LLUVIA las opciones son {' o '.join(CLIMAS)} "
                                f"(llego {valor!r}). No se escribio nada.")

    numero = _a_numero(prioridad, "Prioridad", "1")
    if numero != int(numero) or not PRIORIDAD_MINIMA <= numero <= PRIORIDAD_MAXIMA:
        raise ValorInvalido(f"Prioridad: debe ser un numero entero de {PRIORIDAD_MINIMA} a {PRIORIDAD_MAXIMA} "
                            f"(llego {numero:g}). No se escribio nada.")
    prioridad_n = int(numero)

    mensaje_n = re.sub(r"\s+", " ", _texto(mensaje))
    if not mensaje_n:
        raise ValorInvalido("Mensaje: no puede quedar vacio. No se escribio nada.")
    if len(mensaje_n) > LARGO_MAXIMO_MENSAJE:
        raise ValorInvalido(f"Mensaje: tiene {len(mensaje_n)} caracteres y el maximo es {LARGO_MAXIMO_MENSAJE}. "
                            "No se escribio nada.")
    if mensaje_n[0] in "=+-@":
        raise ValorInvalido(f"Mensaje: no puede empezar con '{mensaje_n[0]}' (el Excel lo tomaria como una formula). "
                            "No se escribio nada.")

    numero = _a_numero(descuento, "Descuento", "10")
    if not 0 <= numero <= DESCUENTO_MAXIMO:
        raise ValorInvalido(f"Descuento: debe estar entre 0 y {DESCUENTO_MAXIMO} % (llego {numero:g}). "
                            "No se escribio nada.")
    descuento_n = _un_decimal(numero, "Descuento")

    return {"tipo": tipo_n, "operador": operador_n, "valor": valor_n, "prioridad": prioridad_n,
            "mensaje": mensaje_n, "descuento": descuento_n}


# ------------------------------------------------------------------ lectura de la hoja
def _columnas(ws):
    """(fila del encabezado, {clave: numero de columna}). Busca el encabezado por su texto."""
    for fila in ws.iter_rows(min_row=1, max_row=15):
        textos = {_texto(c.value): c.column for c in fila if c.value is not None}
        if ENCABEZADOS["nombre"] in textos:
            faltan = [t for t in ENCABEZADOS.values() if t not in textos]
            if faltan:
                raise ValorInvalido(f"En '{HOJA}' faltan los encabezados {faltan}. No se escribio nada.")
            return fila[0].row, {clave: textos[texto] for clave, texto in ENCABEZADOS.items()}
    raise ValorInvalido(f"No encontre el encabezado '{ENCABEZADOS['nombre']}' en '{HOJA}'. No se escribio nada.")


def _filas_de_combos(ws, fila_encabezado, columnas):
    """Filas de la lista de combos: desde el encabezado hasta la primera fila vacia (igual que el catalogo)."""
    filas = []
    for fila in range(fila_encabezado + 1, ws.max_row + 1):
        nombre = _texto(ws.cell(fila, columnas["nombre"]).value)
        codigo = _texto(ws.cell(fila, columnas["codigo"]).value)
        if not nombre and not codigo:
            break
        if nombre and codigo:
            filas.append(fila)
    return filas


def leer_combos(wb_valores=None):
    """Los combos tal como estan en el Excel ya calculado (solo lectura), en el orden de la hoja.

    Cada uno: {'fila', 'codigo', 'nombre', 'detalle', 'precio', 'activo', 'clima', 'tipo', 'operador',
               'valor', 'prioridad', 'mensaje', 'descuento' (en por ciento), 'celdas': {clave: 'G5'},
               'configurable': False si su detonador no es TEMPERATURA ni LLUVIA}
    """
    wb_valores = wb_valores or leer_excel()
    ws = wb_valores[HOJA]
    fila_encabezado, columnas = _columnas(ws)
    combos = []
    for fila in _filas_de_combos(ws, fila_encabezado, columnas):
        leido = {clave: ws.cell(fila, columna).value for clave, columna in columnas.items()}
        tipo = _texto(leido["tipo"]).upper()
        combos.append({
            "fila": fila, "codigo": _texto(leido["codigo"]), "nombre": _texto(leido["nombre"]),
            "detalle": _texto(leido["detalle"]), "precio": _numero(leido["precio"]),
            "activo": _texto(leido["estado"]).upper() == "ACTIVO", "clima": _texto(leido["clima"]),
            "tipo": tipo, "operador": re.sub(r"\s+", "", _texto(leido["operador"])),
            "valor": leido["valor"] if _numero(leido["valor"]) is not None else _texto(leido["valor"]),
            "prioridad": leido["prioridad"], "mensaje": _texto(leido["mensaje"]),
            "descuento": descuento_de_celda(leido["descuento"]),
            "celdas": {clave: ws.cell(fila, columna).coordinate for clave, columna in columnas.items()},
            "configurable": tipo in TIPOS,
        })
    return combos


def _igual(anterior, nuevo):
    if _numero(nuevo) is not None:
        return _numero(anterior) is not None and round(float(anterior), 6) == round(float(nuevo), 6)
    return isinstance(anterior, str) and anterior.strip() == nuevo


# ------------------------------------------------------------------ escritura
def actualizar_combo(codigo, tipo, operador, valor, prioridad, mensaje, descuento, origen="SaaS"):
    """Guarda la configuracion de UN combo (celdas G a L de su fila) y revisa C16.

    codigo: 'co01' (sin distinguir mayusculas). Los demas datos llegan como los digito la persona.
    Devuelve {'hubo_cambios', 'estado', 'codigo', 'nombre', 'fila', 'cambios': [{'campo', 'celda',
              'antes', 'despues'}], 'evento', 'activacion': resultado de la revision de C16}.
    """
    buscado = _texto(codigo).lower()

    def _editar(wb, wb_valores):
        ws = wb[HOJA]
        fila_encabezado, columnas = _columnas(ws)
        filas = [f for f in _filas_de_combos(ws, fila_encabezado, columnas)
                 if _texto(ws.cell(f, columnas["codigo"]).value).lower() == buscado]
        if len(filas) != 1:
            raise ComboDesconocido(f"El combo '{codigo}' aparece {len(filas)} veces en '{HOJA}' "
                                   "(debe aparecer una sola vez). No se escribio nada.")
        fila = filas[0]
        celda = {clave: ws.cell(fila, columna) for clave, columna in columnas.items()}
        nombre = _texto(celda["nombre"].value)
        nuevo = normalizar_configuracion(tipo, operador, valor, prioridad, mensaje, descuento,
                                         operador_actual=celda["operador"].value)
        a_escribir = dict(nuevo, descuento=descuento_a_celda(nuevo["descuento"]))

        cambios = []
        for clave in CAMPOS:
            anterior = celda[clave].value
            if _es_formula(anterior):
                raise ValorInvalido(f"'{HOJA}'!{celda[clave].coordinate} es una formula: el robot no la sobrescribe.")
            if clave == "operador" and nuevo["tipo"] == LLUVIA:
                continue                                  # con LLUVIA no se usa: la celda queda como esta
            if clave == "descuento":
                igual = descuento_de_celda(anterior) == nuevo["descuento"] and _numero(anterior) is not None
            else:
                igual = _igual(anterior, a_escribir[clave])
            if not igual:
                cambios.append({"campo": clave, "celda": celda[clave].coordinate, "antes": anterior,
                                "despues": a_escribir[clave], "titulo": ENCABEZADOS[clave]})
        base = {"codigo": _texto(celda["codigo"].value), "nombre": nombre, "fila": fila, "configuracion": nuevo}
        if not cambios:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", "cambios": [], **base}

        # La columna B (texto de referencia) se alinea con H e I en el mismo guardado.
        criterio = texto_criterio(nuevo["tipo"], nuevo["operador"], nuevo["valor"])
        if _es_formula(celda["clima"].value):
            raise ValorInvalido(f"'{HOJA}'!{celda['clima'].coordinate} es una formula: el robot no la sobrescribe.")
        if celda["clima"].value != criterio:
            cambios.append({"campo": "clima", "celda": celda["clima"].coordinate, "antes": celda["clima"].value,
                            "despues": criterio, "titulo": ENCABEZADOS["clima"]})
        for cambio in cambios:
            celda[cambio["campo"]].value = cambio["despues"]
        detalle = f"{nombre} ({base['codigo']}), fila {fila}: " + "; ".join(
            f"{c['titulo']} {c['celda']}: {c['antes']!r} -> {c['despues']!r}" for c in cambios)
        evento = registrar_evento(
            wb, modulo="Combos", tipo_evento=TIPO_EVENTO, entidad=HOJA, id_entidad=base["codigo"],
            accion="UPDATE", estado="OK", origen=origen, valor_referencia=len(cambios), detalle=detalle,
        )
        return {"hubo_cambios": True, "estado": "OK", "cambios": cambios, "evento": evento, **base}

    resultado = editar_excel(_editar, descripcion=f"Configuracion del combo {codigo}")
    if resultado.get("hubo_cambios"):
        # Con el Excel ya recalculado, el robot revisa 'Activacion de Promocion' (C16, #84).
        resultado["activacion"] = ap.actualizar_activacion_promocion(origen=origen)
    return resultado


if __name__ == "__main__":
    argumentos = [a for a in sys.argv[1:] if a != "--local"]
    if argumentos != ["--ver"]:
        sys.exit("Uso: py -m microservicios.grupo3_integraciones.combos_config [--local] --ver")
    for c in leer_combos():
        criterio = f"{c['tipo']} {c['operador']} {c['valor']}" if c["tipo"] == TEMPERATURA else f"{c['tipo']} = {c['valor']}"
        print(f"  fila {c['fila']}  {c['codigo']}  {c['nombre']:<8}  {'ACTIVO  ' if c['activo'] else 'INACTIVO'}  "
              f"{criterio:<22}  prioridad {c['prioridad']}  descuento {c['descuento']:g} %  "
              f"precio {c['precio']:,.2f}  | {c['mensaje']}")
