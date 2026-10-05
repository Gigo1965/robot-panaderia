"""Microservicio: celdas MANUALES de NODO INTEGRAL (Grupo 3, Decisiones Pendientes #64).

Mapa de celdas 'NODO INTEGRAL' (columna "Modo de Control"): seis filas son "Manual"; las demas
las escribe el robot por horario (nodo_integral.py, competencia.py, pedidos_digitales.py) o son
formulas, y aqui solo se LEEN (Regla #4).

Hallazgo presentado al usuario (4-oct-2026, Decisiones Pendientes #80): de las seis celdas
manuales, solo dos tienen quien las lea:
    C3  Tarifa Base KM       -> 'Interface digital'!B4 y el robot para cobrar el envio (Regla #14)
    C22 Numero SINPE Movil   -> el bot de WhatsApp (catalogo.leer_reglas_del_dia)
C4 (Meta de Venta Diaria), C8 (Accion Meta Ads), C14 (Oferta Fija por Feriado) y C15 (Oferta
Automatica Dia Pago) no las lee ninguna formula ni el codigo; C14 y C15 ademas quedaron sin
efecto por la Regla #8 (en feriado y dia de pago hay descuento general y NO hay combos).

Decisiones del usuario (4-oct-2026, las seis del plan 3a):
  1) C3  Tarifa Base KM: numero mayor que 0, con tope 5,000 (el 0 se rechaza: sin tarifa el
         robot no puede cobrar el envio).
  2) C22 Numero SINPE Movil: ocho digitos, se guarda con el formato 8888-8888 (numero supuesto,
         Regla #27).
  3) C4  Meta de Venta Diaria: de 0 en adelante. Hoy es un dato de referencia.
  4) C8  Accion Meta Ads: lista Subir / Mantener.
  5) C14 y C15: SOLO LECTURA (no se borra nada del Excel, Regla #23).
  6) Las cambia solo la administracion (el SaaS exige ese permiso).
Detalles de construccion aceptados por el usuario (4-oct-2026): la celda se ubica por su
ETIQUETA de la columna B (igual que el bot), no por su numero de fila; C3 se guarda en colones
enteros; C4 tiene un tope de seguridad de 99,999,999 contra un numero mal digitado.

Reglas del proyecto que se respetan:
  * Es CONFIGURACION (como descuento_general.py y existencia_minima.py): el robot actualiza UNA
    celda (UPDATE) y deja un evento en RegistroEventos con el VALOR ANTERIOR -> VALOR NUEVO.
  * Nunca se escribe sobre una formula, ni en una fila que el mapa no marca como manual.
  * Si el valor nuevo es igual al actual, no se sube nada (estado SIN_CAMBIO).
  * Un numero escrito con separador de miles ("150,000" o "1.500") se RECHAZA en vez de
    adivinarlo: se pide escribirlo sin separador (150000).

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo3_integraciones.nodo_manual --local --ver
    py -m microservicios.grupo3_integraciones.nodo_manual --local tarifa_km 550
"""
import math
import os
import re
import sys

# --local: igual que descuento_general.py y las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import editar_excel, leer_excel                 # noqa: E402
from conexion_drive.tablas import registrar_evento                  # noqa: E402

HOJA = "NODO INTEGRAL"
COL_CATEGORIA, COL_ETIQUETA, COL_VALOR, COL_USUARIO, COL_ROBOT = 1, 2, 3, 4, 5      # columnas A a E
FILA_ENCABEZADO = 2
TIPO_EVENTO = "CONFIG_NODO_INTEGRAL"
OPCIONES_META_ADS = ("Subir", "Mantener")

# Las cuatro celdas manuales que se pueden cambiar, en el orden de la hoja.
# 'celda_mapa' es solo documentacion (donde esta hoy segun el mapa): el robot la ubica por la etiqueta.
CAMPOS = {
    "tarifa_km": {
        "etiqueta": "Tarifa Base KM", "celda_mapa": "C3", "tipo": "colones", "minimo": 1, "maximo": 5000,
        "decimales": 0, "titulo": "Tarifa base por kilómetro",
        "nota": "La leen 'Interface digital' B4 y el robot para cobrar el envío (km x tarifa).",
    },
    "meta_venta": {
        "etiqueta": "Meta de Venta Diaria", "celda_mapa": "C4", "tipo": "colones", "minimo": 0, "maximo": 99_999_999,
        "decimales": 2, "titulo": "Meta de venta diaria",
        "nota": "Dato de referencia: hoy ninguna fórmula ni el robot la usan.",
    },
    "meta_ads": {
        "etiqueta": "Acción Meta Ads", "celda_mapa": "C8", "tipo": "lista", "opciones": OPCIONES_META_ADS,
        "titulo": "Acción Meta Ads",
        "nota": "Decisión manual de la administración (subir o mantener la publicidad). Dato de referencia.",
    },
    "sinpe": {
        "etiqueta": "Número SINPE Móvil", "celda_mapa": "C22", "tipo": "telefono",
        "titulo": "Número SINPE Móvil",
        "nota": "El bot de WhatsApp lo muestra al cliente que paga por SINPE Móvil. Número supuesto (Regla #27).",
    },
}

# Filas que el mapa marca "Manual" pero quedan SOLO LECTURA por decision del usuario (#80, decision 5).
SOLO_LECTURA = {
    "Oferta Fija por Feriado": "Sin efecto por la Regla #8: en feriado hay descuento general y no hay combos.",
    "Oferta Automática Día Pago": "Sin efecto por la Regla #8: en día de pago hay descuento general y no hay combos.",
}


class ValorInvalido(Exception):
    """Dato rechazado por una validacion: no se escribio nada."""


class CampoDesconocido(ValorInvalido):
    """El campo pedido no es una de las celdas manuales que se pueden cambiar."""


def _igual(texto_a, texto_b):
    return str(texto_a or "").strip().lower() == str(texto_b or "").strip().lower()


def _numero(valor):
    return valor if isinstance(valor, (int, float)) and not isinstance(valor, bool) else None


def _definicion(campo):
    clave = str(campo or "").strip().lower()
    if clave not in CAMPOS:
        raise CampoDesconocido(
            f"'{campo}' no es una celda manual de '{HOJA}' que se pueda cambiar "
            f"(las que si: {', '.join(CAMPOS)}). No se escribio nada.")
    return clave, CAMPOS[clave]


def _normalizar_colones(valor, definicion):
    nombre = definicion["etiqueta"]
    if isinstance(valor, bool):
        raise ValorInvalido(f"{nombre}: {valor!r} no es un monto.")
    if _numero(valor) is not None:
        numero = float(valor)
    else:
        texto = str(valor if valor is not None else "").replace("₡", "").replace(" ", "").strip()
        if re.fullmatch(r"\d{1,3}([.,]\d{3})+", texto) or ("," in texto and "." in texto):
            raise ValorInvalido(
                f"{nombre}: '{valor}' trae separador de miles. Escriba el monto sin separador, "
                "por ejemplo 150000. No se escribio nada.")
        try:
            numero = float(texto.replace(",", "."))
        except ValueError:
            raise ValorInvalido(f"{nombre}: {valor!r} no es un monto. Escriba un numero, por ejemplo 500.") from None
    if math.isnan(numero) or math.isinf(numero):
        raise ValorInvalido(f"{nombre}: {valor!r} no es un monto. Escriba un numero, por ejemplo 500.")
    numero = round(numero, 2)
    if definicion["decimales"] == 0 and numero != int(numero):
        raise ValorInvalido(f"{nombre}: va en colones enteros, sin centimos (llego {numero:g}). No se escribio nada.")
    numero = round(numero, definicion["decimales"])
    if not definicion["minimo"] <= numero <= definicion["maximo"]:
        raise ValorInvalido(
            f"{nombre}: debe estar entre {definicion['minimo']:,} y {definicion['maximo']:,} colones "
            f"(llego {numero:,.10g}). No se escribio nada.")
    return int(numero) if numero == int(numero) else numero


def _normalizar_lista(valor, definicion):
    for opcion in definicion["opciones"]:
        if _igual(valor, opcion):
            return opcion
    raise ValorInvalido(
        f"{definicion['etiqueta']}: {valor!r} no es una opcion. Las opciones son "
        f"{' o '.join(definicion['opciones'])}. No se escribio nada.")


def _normalizar_telefono(valor, definicion):
    texto = str(valor if valor is not None and not isinstance(valor, bool) else "").strip()
    digitos = re.sub(r"[\s-]", "", texto)
    if not re.fullmatch(r"\d{8}", digitos):
        raise ValorInvalido(
            f"{definicion['etiqueta']}: {valor!r} no es un numero de ocho digitos. "
            "Escribalo como 8888-8888. No se escribio nada.")
    return f"{digitos[:4]}-{digitos[4:]}"


def normalizar_valor(campo, valor):
    """Valida un dato y lo devuelve como se guarda en la celda. Si no se permite, ValorInvalido.

    tarifa_km   '550' -> 550 (colones enteros, de 1 a 5,000)
    meta_venta  '175000' o '175000,50' -> 175000 / 175000.5 (de 0 a 99,999,999)
    meta_ads    'subir' -> 'Subir' (Subir o Mantener)
    sinpe       '70001234' o '7000 1234' -> '7000-1234'
    """
    _, definicion = _definicion(campo)
    if definicion["tipo"] == "colones":
        return _normalizar_colones(valor, definicion)
    if definicion["tipo"] == "lista":
        return _normalizar_lista(valor, definicion)
    return _normalizar_telefono(valor, definicion)


def _fila_de(ws, etiqueta):
    """Numero de fila cuya columna B dice exactamente esa etiqueta (sin distinguir mayusculas)."""
    filas = [f for f in range(1, ws.max_row + 1) if _igual(ws.cell(f, COL_ETIQUETA).value, etiqueta)
             and isinstance(ws.cell(f, COL_ETIQUETA).value, str)]
    if len(filas) != 1:
        raise ValorInvalido(
            f"La etiqueta '{etiqueta}' aparece {len(filas)} veces en la columna B de '{HOJA}' "
            "(debe aparecer una sola vez). No se escribio nada.")
    return filas[0]


def leer_nodo(wb_valores=None):
    """Todas las filas de 'NODO INTEGRAL' tal como estan en el Excel (solo lectura).

    Devuelve una lista de diccionarios en el orden de la hoja:
      {'fila', 'celda', 'categoria', 'descripcion', 'valor', 'accion_usuario', 'accion_robot',
       'clase', 'campo', 'nota', 'definido'}
    clase: 'manual' (se puede cambiar con actualizar_dato_manual), 'sin_efecto' (manual segun
           el mapa, hoy solo lectura), o 'lectura' (la escribe el robot o es una formula).
    definido: False si una celda de lista todavia tiene el texto de relleno del Excel.
    """
    wb_valores = wb_valores or leer_excel()
    ws = wb_valores[HOJA]
    por_etiqueta = {d["etiqueta"].lower(): clave for clave, d in CAMPOS.items()}
    filas, categoria = [], ""
    for fila in range(FILA_ENCABEZADO + 1, ws.max_row + 1):
        etiqueta = ws.cell(fila, COL_ETIQUETA).value
        if not isinstance(etiqueta, str) or not etiqueta.strip():
            continue
        etiqueta = etiqueta.strip()
        categoria = str(ws.cell(fila, COL_CATEGORIA).value or "").strip() or categoria
        valor = ws.cell(fila, COL_VALOR).value
        campo = por_etiqueta.get(etiqueta.lower())
        nota_lectura = next((n for e, n in SOLO_LECTURA.items() if _igual(e, etiqueta)), None)
        definido = True
        if campo and CAMPOS[campo]["tipo"] == "lista":
            definido = any(_igual(valor, o) for o in CAMPOS[campo]["opciones"])
        filas.append({
            "fila": fila, "celda": ws.cell(fila, COL_VALOR).coordinate, "categoria": categoria,
            "descripcion": etiqueta, "valor": valor,
            "accion_usuario": str(ws.cell(fila, COL_USUARIO).value or "").strip(),
            "accion_robot": str(ws.cell(fila, COL_ROBOT).value or "").strip(),
            "clase": "manual" if campo else ("sin_efecto" if nota_lectura else "lectura"),
            "campo": campo, "nota": CAMPOS[campo]["nota"] if campo else (nota_lectura or ""),
            "definido": definido,
        })
    return filas


def actualizar_dato_manual(campo, valor, origen="SaaS"):
    """Cambia UNA celda manual de NODO INTEGRAL.

    campo: 'tarifa_km', 'meta_venta', 'meta_ads' o 'sinpe' (ver CAMPOS).
    valor: el dato tal como lo digito la persona; se valida con normalizar_valor.
    """
    clave, definicion = _definicion(campo)
    nuevo = normalizar_valor(clave, valor)

    def _editar(wb, wb_valores):
        ws = wb[HOJA]
        celda = ws.cell(_fila_de(ws, definicion["etiqueta"]), COL_VALOR)
        anterior = celda.value
        if (isinstance(anterior, str) and anterior.startswith("=")) or type(anterior).__name__ == "ArrayFormula":
            raise ValorInvalido(f"'{HOJA}'!{celda.coordinate} es una formula: el robot no la sobrescribe.")
        if _numero(nuevo) is not None:
            igual = _numero(anterior) is not None and round(anterior, 2) == nuevo
        else:
            igual = isinstance(anterior, str) and anterior.strip() == nuevo
        if igual:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", "campo": clave, "celda": celda.coordinate,
                    "etiqueta": definicion["etiqueta"], "valor": nuevo}
        celda.value = nuevo
        evento = registrar_evento(
            wb, modulo="NODO INTEGRAL", tipo_evento=TIPO_EVENTO, entidad=HOJA, id_entidad=celda.coordinate,
            accion="UPDATE", estado="OK", origen=origen, valor_referencia=nuevo,
            detalle=f"'{HOJA}'!{celda.coordinate} {definicion['etiqueta']}: {anterior!r} -> {nuevo!r}",
        )
        return {"hubo_cambios": True, "estado": "OK", "campo": clave, "celda": celda.coordinate, "evento": evento,
                "etiqueta": definicion["etiqueta"], "antes": anterior, "despues": nuevo, "valor": nuevo}

    return editar_excel(_editar, descripcion=f"NODO INTEGRAL {definicion['etiqueta']}: {nuevo}")


if __name__ == "__main__":
    argumentos = [a for a in sys.argv[1:] if a != "--local"]
    if argumentos == ["--ver"]:
        for f in leer_nodo():
            marca = {"manual": "MANUAL    ", "sin_efecto": "SIN EFECTO", "lectura": "lectura   "}[f["clase"]]
            campo = f" [{f['campo']}]" if f["campo"] else ""
            print(f"  {f['celda']:>4}  {marca}  {f['descripcion'][:34]:<34}  {str(f['valor'])[:48]}{campo}")
        sys.exit(0)
    if len(argumentos) != 2:
        sys.exit("Uso: py -m microservicios.grupo3_integraciones.nodo_manual [--local] --ver\n"
                 "     py -m microservicios.grupo3_integraciones.nodo_manual [--local] <campo> <valor>\n"
                 f"     campos: {', '.join(CAMPOS)}")
    try:
        print(actualizar_dato_manual(argumentos[0], argumentos[1], origen="Terminal"))
    except ValorInvalido as error:
        sys.exit(f"RECHAZADO: {error}")
