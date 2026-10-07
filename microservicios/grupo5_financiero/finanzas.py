"""Microservicio: Financiero (Grupo 5) - los datos digitados a mano de las hojas de flujo de caja.

El libro de mapas de celdas es la guia. En las hojas de flujo de caja marca DOS datos manuales;
todo lo demas son formulas que el robot solo LEE:

  registrar_caja_chica(fecha, monto, motivo)   -> 'Flujo caja diario' columna G de la fila de ese dia
      Mapa 'Flujo caja diario': "G5 (patron filas 5 a 369) Gastos / vales / caja chica del dia.
      Input digitado por el usuario final EN EL SaaS -> el ROBOT escribe aqui. UNICA columna donde
      el robot ESCRIBE en esta hoja" (Decisiones Pendientes #66).

  registrar_salida_efectivo(mes, monto)        -> 'flujo caja anual' fila "Salidas de efectivo" (B19:M19)
      Mapa 'Mapa_Flujo Caja Anual': "La UNICA EXCEPCION es 'Salidas de efectivo' (B19:M19): una
      persona lo digita en la interfaz del SaaS, el robot lo recibe, y el robot lo escribe como VALOR
      FIJO en la celda del mes correspondiente (nunca formula viva)" (Decisiones Pendientes #67).

|| 6-oct-2026 (paquete 2): los cuatro SUPUESTOS FINANCIEROS que el mapa marca como manuales:

  actualizar_supuesto("tasa", 12.51)     -> 'flujo caja proyecto' C28 (tasa de interes del prestamo)
  actualizar_supuesto("plazo", 60)       -> 'flujo caja proyecto' B31 (plazo del prestamo en meses)
      Mapa 'flujo caja proyectado', seccion 1: "Usuario (persona): 12.51 (valor fijo digitado) -> C28;
      60 (valor fijo digitado) -> B31" (Decisiones Pendientes #68). El capital de trabajo B37:B40 ya
      son formulas (seccion 8 del mapa): no se digita.
  actualizar_supuesto("mezcla", 50)      -> 'Punto de Equilibrio' B11 (mezcla de ventas de Baguette)
  actualizar_supuesto("utilidad", 20)    -> 'Punto de Equilibrio' B22 (% de utilidad esperada anual)
      Mapa 'Punto de Equilibrio' y 'Diseño del robot (6)': "DOS celdas de input (B11 y B22)...
      el ADMINISTRADOR define o ajusta estos 2 valores en una pantalla del SaaS; el robot los recibe y
      los escribe aqui como valor fijo (UPDATE permitido)" (Decisiones Pendientes #69).

  Decisiones del usuario (6-oct-2026) para estos cuatro datos:
  * Cada celda guarda el numero como ya lo guarda hoy: C28 guarda 12.51 (sus formulas lo dividen
    entre 100: B30 y 'TIR VAR' C11); B11 y B22 guardan la fraccion (50 % = 0.5; 20 % = 0.2). En la
    pantalla siempre se digita el porcentaje.
  * Los textos fijos 'flujo caja proyecto' A29 ("taza interes 12.51%") y 'Punto de Equilibrio' A23
    ("... del 20%") NO se tocan: el mapa no los marca como celdas que escriba el robot.
  * Solo se rechaza lo que romperia las formulas: un dato que no es numero, un porcentaje fuera de
    0 a 100 o un plazo que no es un numero entero de meses desde 1. Si un cambio dejara un error de
    formula (por ejemplo una division entre cero), el robot no sube nada: es la regla general de
    editar_excel para cualquier escritura.
  * Plazo distinto de 60 meses: se guarda con AVISO (la fila 21 resta la cuota anual en los cinco años).
  * Margen negativo (una hornada pequeña cargada en la hoja de costo sube el costo unitario sobre el
    precio): NO se bloquea; es un indicador de que la operacion deja perdida. Se guarda con AVISO.
  Lecturas para el SaaS (no escriben nada): leer_flujo_proyecto, leer_punto_equilibrio, leer_tir_var,
  leer_proyeccion_ventas, leer_proyeccion_anual y leer_indicadores (VAN, TIR, punto de equilibrio).

Decisiones del usuario (6-oct-2026):
  * Se sigue el mapa de celdas en todo. Las filas vacias sin funcion (por ejemplo la fila 26 de
    'flujo caja anual') son espacios de orden: el robot no las toca.
  * Caja chica: se digita el TOTAL DEL DIA con su motivo; un total nuevo REEMPLAZA al anterior
    (el mapa habla de un "cierre puntual para el ingreso manual de caja chica"). El evento guarda
    el valor anterior, el nuevo y el motivo: no se pierde informacion.
  * Solo se aceptan fechas que ya pasaron o la de hoy (la hoja no genera fechas futuras) y montos
    de 0 a 500,000 colones.
  * Las pantallas del Grupo 5 son solo para administracion (el permiso lo exige el SaaS).

Como se ubica el dia: la columna A de 'Flujo caja diario' es una formula (A5 = la primera fecha con
movimiento; cada fila siguiente = la anterior + 1, hasta HOY). La fila de una fecha es entonces
5 + (fecha - primera fecha). El robot lee la primera fecha ya calculada, hace esa cuenta y, si la
fila ya trae su fecha calculada, comprueba que sea la misma antes de escribir.

Reglas del proyecto que se respetan:
  * Es un valor que digita una persona: el robot escribe UNA celda (UPDATE) y deja un evento en
    RegistroEventos con el VALOR ANTERIOR -> VALOR NUEVO.
  * Nunca se escribe sobre una formula. Las celdas se ubican por el texto de su encabezado o de su
    etiqueta, no por su numero de fila o su letra.
  * Si el valor nuevo es igual al actual, no se sube nada (estado SIN_CAMBIO).
  * Un numero con separador de miles ("150,000") se RECHAZA en vez de adivinarlo.

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo5_financiero.finanzas --local --ver
    py -m microservicios.grupo5_financiero.finanzas --local caja-chica 2026-10-06 3500 "Vales del dia"
    py -m microservicios.grupo5_financiero.finanzas --local salidas octubre 25000
    py -m microservicios.grupo5_financiero.finanzas --local supuesto tasa 12.51
"""
import math
import os
import re
import sys
from datetime import date, datetime, timedelta

# --local: igual que costeo.py y las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import editar_excel, leer_excel                               # noqa: E402
from conexion_drive.tablas import columnas_de_tabla, filas_de_datos, obtener_tabla, registrar_evento   # noqa: E402
from conexion_drive.tiempo import ahora_local                                     # noqa: E402

MODULO = "Financiero"
ORIGEN_ROBOT = "Robot Python"              # origen del evento cuando nadie entrega otro (#78)

# --- 'Flujo caja diario' (mapa 'Flujo caja diario'; tabla tblFlujoCajaDiario) ---
HOJA_DIARIO = "Flujo caja diario"
TABLA_DIARIO = "tblFlujoCajaDiario"
COLUMNA_FECHA = "Fecha"
TEXTO_CAJA_CHICA = "caja chica"            # el encabezado de la columna G: 'Gastos / vales / caja chica'
EVENTO_CAJA_CHICA = "CAJA_CHICA"
TOPE_CAJA_CHICA = 500_000                  # decision del usuario (6-oct-2026): de 0 a 500,000
LARGO_MOTIVO = 120
DIAS_EN_PANTALLA = 14
# Encabezado de la tabla -> nombre corto con el que leer_flujo_diario entrega cada columna.
COLUMNAS_DIARIO = {
    "ventas despacho": "ventas_despacho", "ventas digitales": "ventas_digitales",
    "envíos cobrados": "envios", "total ingresos día": "ingresos", "materia prima": "materia_prima",
    "compras externas netas": "compras_externas", "gastos / vales / caja chica": "caja_chica",
    "total egresos día": "egresos", "flujo neto día": "neto", "saldo acumulado": "saldo",
}

# --- 'flujo caja anual' (mapa 'Mapa_Flujo Caja Anual') ---
HOJA_ANUAL = "flujo caja anual"
ETIQUETA_PERIODO = "periodo"               # A5: la fila de los nombres de los meses
ETIQUETA_SALIDAS = "salidas de efectivo"   # A19: la unica fila manual de la hoja segun el mapa
EVENTO_SALIDAS = "SALIDAS_EFECTIVO"
TOPE_SALIDAS = 99_999_999                  # tope de seguridad contra un numero mal digitado
COLUMNAS_MES = "BCDEFGHIJKLM"              # enero = B ... diciembre = M
COLUMNA_TOTAL = "N"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


# --- Supuestos financieros: 'flujo caja proyecto' y 'Punto de Equilibrio' (paquete 2) ---
HOJA_PROYECTO = "flujo caja proyecto"
HOJA_EQUILIBRIO = "Punto de Equilibrio"
HOJA_TIR = "TIR VAR"
HOJA_PROYECCION = "proyeccion ventas"
HOJA_PROYECCION_ANUAL = "proyeccion anual ventas"
EVENTO_PRESTAMO = "CONFIG_PRESTAMO"
EVENTO_EQUILIBRIO = "CONFIG_PUNTO_EQUILIBRIO"
PLAZO_DEL_MODELO = 60                      # la proyeccion resta la cuota en los cinco años (fila 21)
TOPE_PLAZO = 600                           # tope de seguridad contra un numero mal digitado (50 años)
# 'celda_mapa' es solo documentacion (donde esta hoy segun el mapa): el robot ubica la celda por la
# etiqueta de la columna A de su fila. 'guarda': como queda el numero en la celda.
SUPUESTOS = {
    "tasa": {"hoja": HOJA_PROYECTO, "celda_mapa": "C28", "columna": "C", "tipo": "porcentaje", "guarda": "numero",
             "etiqueta": "Tasa de interes del prestamo", "titulo": "Tasa de interés del préstamo",
             "debe_decir": ("banco", "prestamo"), "evento": EVENTO_PRESTAMO},
    "plazo": {"hoja": HOJA_PROYECTO, "celda_mapa": "B31", "columna": "B", "tipo": "meses", "guarda": "numero",
              "etiqueta": "Plazo del prestamo (meses)", "titulo": "Plazo del préstamo",
              "debe_decir": ("plazo",), "evento": EVENTO_PRESTAMO},
    "mezcla": {"hoja": HOJA_EQUILIBRIO, "celda_mapa": "B11", "columna": "B", "tipo": "porcentaje", "guarda": "fraccion",
               "etiqueta": "Mezcla de ventas Baguette", "titulo": "Mezcla de ventas de Baguette",
               "debe_decir": ("mezcla de ventas", "baguette"), "evento": EVENTO_EQUILIBRIO},
    "utilidad": {"hoja": HOJA_EQUILIBRIO, "celda_mapa": "B22", "columna": "B", "tipo": "porcentaje", "guarda": "fraccion",
                 "etiqueta": "Utilidad esperada anual", "titulo": "Utilidad esperada anual",
                 "debe_decir": ("% utilidad esperada",), "evento": EVENTO_EQUILIBRIO},
}


class FinanzasInvalido(Exception):
    """Dato rechazado por una validacion: no se escribio nada."""


# ----------------------------------------------------------------------------
# Validaciones
# ----------------------------------------------------------------------------
def _clave(texto):
    return " ".join(str(texto or "").split()).casefold()


def _numero(valor):
    return valor if isinstance(valor, (int, float)) and not isinstance(valor, bool) else None


def _es_formula(valor):
    return (isinstance(valor, str) and valor.startswith("=")) or type(valor).__name__ == "ArrayFormula"


def _monto(valor, nombre, maximo):
    """Monto en colones, de 0 a maximo, con 2 decimales. '3500', '3500.50' o '3500,50'."""
    if isinstance(valor, bool):
        raise FinanzasInvalido(f"{nombre}: {valor!r} no es un monto. No se escribio nada.")
    if _numero(valor) is not None:
        numero = float(valor)
    else:
        texto = str(valor if valor is not None else "").replace("₡", "").replace(" ", "").strip()
        if re.fullmatch(r"\d{1,3}([.,]\d{3})+", texto) or ("," in texto and "." in texto):
            raise FinanzasInvalido(
                f"{nombre}: '{valor}' trae separador de miles. Escriba el monto sin separador, "
                "por ejemplo 15000. No se escribio nada.")
        try:
            numero = float(texto.replace(",", "."))
        except ValueError:
            raise FinanzasInvalido(
                f"{nombre}: {valor!r} no es un monto. Escriba un numero, por ejemplo 3500. "
                "No se escribio nada.") from None
    if math.isnan(numero) or math.isinf(numero):
        raise FinanzasInvalido(f"{nombre}: {valor!r} no es un monto. No se escribio nada.")
    numero = round(numero, 2)
    if not 0 <= numero <= maximo:
        raise FinanzasInvalido(
            f"{nombre}: debe estar entre 0 y {maximo:,} colones (llego {numero:,.10g}). No se escribio nada.")
    return int(numero) if numero == int(numero) else numero


def _fecha(valor, hoy):
    """La fecha del dia: un date, '2026-10-06' o '06/10/2026'. Vacio = hoy."""
    if valor in (None, ""):
        return hoy
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor).strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue
    raise FinanzasInvalido(
        f"Fecha: {valor!r} no es una fecha. Escribala como 2026-10-06 o 06/10/2026. No se escribio nada.")


def _motivo(valor):
    texto = " ".join(str(valor if valor is not None else "").split())
    if not texto:
        raise FinanzasInvalido("Motivo: escriba en que se gasto (por ejemplo 'Vales del dia'). No se escribio nada.")
    if texto[0] in "=+-@":
        raise FinanzasInvalido("Motivo: no puede empezar con '=', '+', '-' ni '@' (regla del proyecto). "
                               "No se escribio nada.")
    if len(texto) > LARGO_MOTIVO:
        raise FinanzasInvalido(f"Motivo: maximo {LARGO_MOTIVO} letras (llegaron {len(texto)}). No se escribio nada.")
    return texto


def _mes(valor):
    """1 a 12, o el nombre del mes ('octubre'). Devuelve el numero."""
    if isinstance(valor, str) and _clave(valor) in MESES:
        return MESES.index(_clave(valor)) + 1
    if _clave(valor) == "setiembre":
        return 9
    try:
        numero = float(str(valor).strip())
    except (TypeError, ValueError):
        numero = None
    if isinstance(valor, bool) or numero is None or numero != int(numero) or not 1 <= numero <= 12:
        raise FinanzasInvalido(f"Mes: debe ser de 1 a 12 o el nombre del mes (llego {valor!r}). No se escribio nada.")
    return int(numero)


# ----------------------------------------------------------------------------
# 'Flujo caja diario': donde esta cada cosa (por el texto del encabezado)
# ----------------------------------------------------------------------------
def _plano_diario(wb):
    """Las letras de las columnas de la tabla y sus filas, leidas de la tabla misma."""
    ws, tabla = obtener_tabla(wb, HOJA_DIARIO, TABLA_DIARIO)
    columnas = columnas_de_tabla(tabla)
    por_clave = {_clave(nombre): letra for nombre, letra in columnas.items()}
    fecha = por_clave.get(_clave(COLUMNA_FECHA))
    caja = [letra for clave, letra in por_clave.items() if TEXTO_CAJA_CHICA in clave]
    if fecha is None or len(caja) != 1:
        raise KeyError(f"La tabla {TABLA_DIARIO} de '{HOJA_DIARIO}' no tiene la columna '{COLUMNA_FECHA}' o la de "
                       "caja chica: el mapa y el Excel no coinciden.")
    filas = filas_de_datos(ws, tabla)
    return {"ws": ws, "fecha": fecha, "caja": caja[0], "primera_fila": filas[0], "ultima_fila": filas[-1],
            "por_clave": por_clave}


def _primera_fecha(ws_valores, plano):
    valor = ws_valores[f"{plano['fecha']}{plano['primera_fila']}"].value
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    raise KeyError(f"'{HOJA_DIARIO}'!{plano['fecha']}{plano['primera_fila']} no trae una fecha calculada "
                   f"(trae {valor!r}): no se puede ubicar el dia.")


def _fila_del_dia(ws_valores, plano, fecha, hoy):
    """Numero de fila de esa fecha. Rechaza fechas futuras, anteriores a la hoja o fuera de la tabla."""
    primera = _primera_fecha(ws_valores, plano)
    if fecha > hoy:
        raise FinanzasInvalido(
            f"Fecha: {fecha.isoformat()} todavia no llega (hoy es {hoy.isoformat()}). La hoja no lleva fechas "
            "futuras. No se escribio nada.")
    if fecha < primera:
        raise FinanzasInvalido(
            f"Fecha: {fecha.isoformat()} es anterior al primer dia de '{HOJA_DIARIO}' ({primera.isoformat()}). "
            "No se escribio nada.")
    fila = plano["primera_fila"] + (fecha - primera).days
    if fila > plano["ultima_fila"]:
        raise FinanzasInvalido(
            f"Fecha: {fecha.isoformat()} queda fuera de la tabla de '{HOJA_DIARIO}' (termina en la fila "
            f"{plano['ultima_fila']}). No se escribio nada.")
    calculada = ws_valores[f"{plano['fecha']}{fila}"].value
    if isinstance(calculada, datetime):
        calculada = calculada.date()
    if isinstance(calculada, date) and calculada != fecha:
        raise FinanzasInvalido(
            f"'{HOJA_DIARIO}'!{plano['fecha']}{fila} trae la fecha {calculada.isoformat()} y se esperaba "
            f"{fecha.isoformat()}: el mapa y el Excel no coinciden. No se escribio nada.")
    return fila


# ----------------------------------------------------------------------------
# 1) Caja chica del dia ('Flujo caja diario' columna G, #66)
# ----------------------------------------------------------------------------
def registrar_caja_chica(fecha, monto, motivo, origen=ORIGEN_ROBOT, hoy=None):
    """Escribe el TOTAL de gastos / vales / caja chica de un dia. Un total nuevo reemplaza al anterior.

    fecha:  el dia (date, '2026-10-06' o '06/10/2026'); vacio = hoy. No se aceptan fechas futuras.
    monto:  total del dia en colones, de 0 a 500,000 (sin separador de miles).
    motivo: en que se gasto; queda en el evento junto con el valor anterior y el nuevo.
    hoy:    solo para las pruebas (por defecto, la fecha de Costa Rica).
    """
    hoy = hoy or ahora_local().date()
    dia = _fecha(fecha, hoy)
    nuevo = _monto(monto, "Caja chica", TOPE_CAJA_CHICA)
    texto = _motivo(motivo)

    def _editar(wb, wb_valores):
        plano = _plano_diario(wb)
        fila = _fila_del_dia(wb_valores[HOJA_DIARIO], plano, dia, hoy)
        celda = plano["ws"][f"{plano['caja']}{fila}"]
        anterior = celda.value
        if _es_formula(anterior):
            raise FinanzasInvalido(f"'{HOJA_DIARIO}'!{celda.coordinate} es una formula: el robot no la sobrescribe.")
        if anterior not in (None, "") and _numero(anterior) is None:
            raise FinanzasInvalido(f"'{HOJA_DIARIO}'!{celda.coordinate} trae {anterior!r}, que no es un monto: "
                                   "el robot no lo sobrescribe.")
        antes = round(float(anterior or 0), 2)
        base = {"hoja": HOJA_DIARIO, "celda": celda.coordinate, "fila": fila, "fecha": dia, "motivo": texto,
                "antes": antes, "despues": nuevo}
        if antes == nuevo:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", **base}
        celda.value = nuevo
        evento = registrar_evento(
            wb, modulo=MODULO, tipo_evento=EVENTO_CAJA_CHICA, entidad=HOJA_DIARIO, id_entidad=dia.isoformat(),
            accion="UPDATE", estado="OK", origen=origen or ORIGEN_ROBOT, valor_referencia=nuevo,
            detalle=(f"'{HOJA_DIARIO}'!{celda.coordinate} Gastos / vales / caja chica del {dia.isoformat()}: "
                     f"{antes:g} -> {nuevo:g} | motivo: {texto}"))
        return {"hubo_cambios": True, "estado": "OK", "evento": evento, **base}

    return editar_excel(_editar, descripcion=f"Caja chica {dia.isoformat()}: {nuevo}")


def leer_flujo_diario(wb, hoy=None, dias=DIAS_EN_PANTALLA):
    """La hoja 'Flujo caja diario' tal como la calculo el Excel (solo LECTURA; wb = libro de VALORES).

    Devuelve {'primera', 'hoy', 'ultima_calculada', 'saldo', 'resumen': [{'concepto', 'total'}],
              'dias': [...], 'celda_saldo', 'columna_caja', 'tope'}.
    'dias' trae los ultimos dias (el mas reciente primero). Un dia que el Excel todavia no ha
    calculado (nadie ha escrito hoy) viene con 'calculado': False y sus montos en None.
    """
    hoy = hoy or ahora_local().date()
    plano = _plano_diario(wb)
    ws = plano["ws"]
    primera = _primera_fecha(ws, plano)
    ultimo_dia = min(hoy, primera + timedelta(days=plano["ultima_fila"] - plano["primera_fila"]))

    filas = []
    for atras in range(max(dias, 1)):
        dia = ultimo_dia - timedelta(days=atras)
        if dia < primera:
            break
        fila = plano["primera_fila"] + (dia - primera).days
        fecha_calculada = ws[f"{plano['fecha']}{fila}"].value
        calculado = isinstance(fecha_calculada, (date, datetime))
        registro = {"fila": fila, "fecha": dia, "calculado": calculado, "es_hoy": dia == hoy,
                    "celda_caja": f"{plano['caja']}{fila}"}
        for encabezado, corto in COLUMNAS_DIARIO.items():
            letra = plano["por_clave"].get(_clave(encabezado))
            valor = _numero(ws[f"{letra}{fila}"].value) if letra else None
            registro[corto] = valor if calculado or corto == "caja_chica" else None
        filas.append(registro)

    ultima_calculada = None
    for fila in range(plano["ultima_fila"], plano["primera_fila"] - 1, -1):
        valor = ws[f"{plano['fecha']}{fila}"].value
        if isinstance(valor, (date, datetime)):
            ultima_calculada = valor.date() if isinstance(valor, datetime) else valor
            break

    # Bloque "RESUMEN / TOTALES" a la derecha de la tabla: se ubica por el encabezado 'Concepto'.
    resumen, saldo, celda_saldo = [], None, None
    fila_encabezado = plano["primera_fila"] - 1
    columna = next((c for c in range(1, ws.max_column + 1)
                    if _clave(ws.cell(fila_encabezado, c).value) == "concepto"), None)
    if columna:
        for fila in range(fila_encabezado + 1, fila_encabezado + 30):
            concepto = " ".join(str(ws.cell(fila, columna).value or "").split())
            if not concepto:
                break
            total = _numero(ws.cell(fila, columna + 1).value)
            resumen.append({"concepto": concepto, "total": total, "celda": ws.cell(fila, columna + 1).coordinate})
            if "saldo de caja" in concepto.casefold():
                saldo, celda_saldo = total, ws.cell(fila, columna + 1).coordinate
    return {"primera": primera, "hoy": hoy, "ultimo_dia": ultimo_dia, "ultima_calculada": ultima_calculada,
            "dias": filas, "resumen": resumen, "saldo": saldo, "celda_saldo": celda_saldo,
            "columna_caja": plano["caja"], "tope": TOPE_CAJA_CHICA}


# ----------------------------------------------------------------------------
# 2) Salidas de efectivo del mes ('flujo caja anual' fila 19, #67)
# ----------------------------------------------------------------------------
def _plano_anual(ws):
    """Fila de los meses (A = 'Periodo') y fila de 'Salidas de efectivo', ubicadas por su texto."""
    def filas_que_dicen(texto):
        return [f for f in range(1, ws.max_row + 1) if _clave(ws.cell(f, 1).value) == texto]
    periodo, salidas = filas_que_dicen(ETIQUETA_PERIODO), filas_que_dicen(ETIQUETA_SALIDAS)
    if len(periodo) != 1 or len(salidas) != 1:
        raise KeyError(f"En la columna A de '{HOJA_ANUAL}' deben aparecer una sola vez 'Periodo' "
                       f"({len(periodo)}) y 'Salidas de efectivo' ({len(salidas)}): el mapa y el Excel no coinciden.")
    for numero, letra in enumerate(COLUMNAS_MES):
        encabezado = _clave(ws[f"{letra}{periodo[0]}"].value)
        if encabezado not in (MESES[numero], "setiembre" if numero == 8 else MESES[numero]):
            raise KeyError(f"'{HOJA_ANUAL}'!{letra}{periodo[0]} dice {encabezado!r} y se esperaba {MESES[numero]!r}: "
                           "el mapa y el Excel no coinciden.")
    return {"fila_meses": periodo[0], "fila_salidas": salidas[0]}


def registrar_salida_efectivo(mes, monto, origen=ORIGEN_ROBOT):
    """Escribe las 'Salidas de efectivo' de un mes como VALOR FIJO (un monto por mes, B19:M19).

    mes:   1 a 12 o el nombre del mes.
    monto: colones, de 0 en adelante (sin separador de miles).
    """
    numero_mes = _mes(mes)
    nuevo = _monto(monto, "Salidas de efectivo", TOPE_SALIDAS)

    def _editar(wb, wb_valores):
        ws = wb[HOJA_ANUAL]
        plano = _plano_anual(ws)
        celda = ws[f"{COLUMNAS_MES[numero_mes - 1]}{plano['fila_salidas']}"]
        anterior = celda.value
        if _es_formula(anterior):
            raise FinanzasInvalido(f"'{HOJA_ANUAL}'!{celda.coordinate} es una formula: el robot no la sobrescribe.")
        if anterior not in (None, "") and _numero(anterior) is None:
            raise FinanzasInvalido(f"'{HOJA_ANUAL}'!{celda.coordinate} trae {anterior!r}, que no es un monto: "
                                   "el robot no lo sobrescribe.")
        antes = round(float(anterior or 0), 2)
        base = {"hoja": HOJA_ANUAL, "celda": celda.coordinate, "mes": numero_mes, "nombre_mes": MESES[numero_mes - 1],
                "antes": antes, "despues": nuevo}
        if antes == nuevo:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", **base}
        celda.value = nuevo
        vacia = "vacia" if anterior in (None, "") else f"{antes:g}"
        evento = registrar_evento(
            wb, modulo=MODULO, tipo_evento=EVENTO_SALIDAS, entidad=HOJA_ANUAL, id_entidad=MESES[numero_mes - 1],
            accion="UPDATE", estado="OK", origen=origen or ORIGEN_ROBOT, valor_referencia=nuevo,
            detalle=f"'{HOJA_ANUAL}'!{celda.coordinate} Salidas de efectivo de {MESES[numero_mes - 1]}: {vacia} -> {nuevo:g}")
        return {"hubo_cambios": True, "estado": "OK", "evento": evento, **base}

    return editar_excel(_editar, descripcion=f"Salidas de efectivo mes {numero_mes}: {nuevo}")


def leer_flujo_anual(wb):
    """La hoja 'flujo caja anual' tal como la calculo el Excel (solo LECTURA; wb = libro de VALORES).

    Devuelve {'meses': [12 nombres], 'filas': [...], 'salidas': la fila manual}.
    Cada fila: {'fila', 'etiqueta', 'meses': [12 valores], 'total', 'clase'}; clase es 'manual'
    (Salidas de efectivo), 'titulo' (titulo de seccion, sin datos por diseño) o 'formula'.
    Las filas sin etiqueta son espacios de orden y no se devuelven.
    """
    ws = wb[HOJA_ANUAL]
    plano = _plano_anual(ws)
    filas = []
    for fila in range(plano["fila_meses"] + 1, ws.max_row + 1):
        etiqueta = " ".join(str(ws.cell(fila, 1).value or "").split())
        if not etiqueta:
            continue
        meses = [_numero(ws[f"{letra}{fila}"].value) for letra in COLUMNAS_MES]
        total = _numero(ws[f"{COLUMNA_TOTAL}{fila}"].value)
        if fila == plano["fila_salidas"]:
            clase = "manual"
        elif total is None and all(valor is None for valor in meses):
            clase = "titulo"
        else:
            clase = "formula"
        filas.append({"fila": fila, "etiqueta": etiqueta, "meses": meses, "total": total, "clase": clase})
    salidas = next(f for f in filas if f["clase"] == "manual")
    return {"meses": list(MESES), "filas": filas, "salidas": salidas, "tope": TOPE_SALIDAS}


# ----------------------------------------------------------------------------
# 3) Supuestos financieros ('flujo caja proyecto' C28 y B31; 'Punto de Equilibrio' B11 y B22; #68 y #69)
# ----------------------------------------------------------------------------
def _sin_acentos(texto):
    return _clave(texto).translate(str.maketrans("áéíóúü", "aeiouu"))


def _filas_que_dicen(ws, palabras, desde=1, empieza=False):
    """Filas cuya columna A contiene TODAS esas palabras (sin acentos ni mayusculas)."""
    filas = []
    for fila in range(desde, ws.max_row + 1):
        texto = _sin_acentos(ws.cell(fila, 1).value)
        if texto and all(p in texto for p in palabras) and (not empieza or texto.startswith(palabras[0])):
            filas.append(fila)
    return filas


def _celda_supuesto(ws, campo):
    """La celda del supuesto, ubicada por el texto de la columna A de su fila (debe haber una sola)."""
    definicion = SUPUESTOS[campo]
    filas = _filas_que_dicen(ws, definicion["debe_decir"], empieza=campo in ("plazo", "mezcla", "utilidad"))
    if len(filas) != 1:
        raise KeyError(f"En la columna A de '{definicion['hoja']}' debe haber una sola fila que diga "
                       f"{' y '.join(definicion['debe_decir'])} (hay {len(filas)}): el mapa y el Excel no coinciden.")
    return ws[f"{definicion['columna']}{filas[0]}"]


def _definicion_supuesto(campo):
    clave = _clave(campo)
    if clave not in SUPUESTOS:
        raise FinanzasInvalido(f"'{campo}' no es un supuesto financiero que se pueda cambiar "
                               f"(los que si: {', '.join(SUPUESTOS)}). No se escribio nada.")
    return clave, SUPUESTOS[clave]


def _porcentaje(valor, nombre):
    """'12.51', '12,51' o '12.51 %' -> 12.51. De 0 a 100, con 2 decimales."""
    if isinstance(valor, bool):
        raise FinanzasInvalido(f"{nombre}: {valor!r} no es un porcentaje. No se escribio nada.")
    if _numero(valor) is not None:
        numero = float(valor)
    else:
        texto = str(valor if valor is not None else "").replace("%", "").replace(" ", "").strip()
        try:
            numero = float(texto.replace(",", "."))
        except ValueError:
            raise FinanzasInvalido(f"{nombre}: {valor!r} no es un porcentaje. Escriba un numero, por ejemplo "
                                   "12.51. No se escribio nada.") from None
    if math.isnan(numero) or math.isinf(numero) or not 0 <= round(numero, 2) <= 100:
        raise FinanzasInvalido(f"{nombre}: debe estar entre 0 y 100 % (llego {valor!r}). No se escribio nada.")
    numero = round(numero, 2)
    return int(numero) if numero == int(numero) else numero


def _meses(valor, nombre):
    try:
        numero = float(str(valor).replace(" ", "").strip())
    except (TypeError, ValueError):
        numero = None
    if isinstance(valor, bool) or numero is None or math.isnan(numero) or math.isinf(numero) \
            or numero != int(numero) or not 1 <= numero <= TOPE_PLAZO:
        raise FinanzasInvalido(f"{nombre}: debe ser un numero entero de meses, de 1 a {TOPE_PLAZO} "
                               f"(llego {valor!r}). No se escribio nada.")
    return int(numero)


def normalizar_supuesto(campo, valor):
    """Valida el dato y devuelve (lo que se guarda en la celda, el dato como lo ve la persona).

    tasa      '14' -> (14, 14)         la celda guarda el porcentaje tal cual
    plazo     '60' -> (60, 60)         meses enteros
    mezcla    '60' -> (0.6, 60)        la celda guarda la fraccion
    utilidad  '25' -> (0.25, 25)
    """
    clave, definicion = _definicion_supuesto(campo)
    if definicion["tipo"] == "meses":
        meses = _meses(valor, definicion["titulo"])
        return meses, meses
    porcentaje = _porcentaje(valor, definicion["titulo"])
    if definicion["guarda"] == "fraccion":
        return round(porcentaje / 100, 4), porcentaje
    return porcentaje, porcentaje


def _como_se_ve(campo, guardado):
    """El valor de la celda como lo lee una persona: 0.5 -> '50 %', 12.51 -> '12.51 %', 60 -> '60 meses'."""
    definicion = SUPUESTOS[campo]
    if _numero(guardado) is None:
        return repr(guardado)
    if definicion["tipo"] == "meses":
        return f"{guardado:g} meses"
    return f"{round(guardado * 100, 2):g} %" if definicion["guarda"] == "fraccion" else f"{guardado:g} %"


def leer_indicadores(wb):
    """Los resultados que mueven los cuatro supuestos, tal como los calculo el Excel (wb = libro de VALORES)."""
    equilibrio = leer_punto_equilibrio(wb)
    prestamo = leer_flujo_proyecto(wb)["prestamo"]
    tir = leer_tir_var(wb)
    por_celda = {fila["celda"]: fila["valor"] for fila in equilibrio["filas"]}

    def de(palabras, empieza=True):
        filas = _filas_que_dicen(wb[HOJA_EQUILIBRIO], palabras, empieza=empieza)
        return por_celda.get(f"B{filas[0]}") if filas else None
    return {
        "tasa": prestamo["tasa"], "plazo": prestamo["plazo"], "cuota_mensual": prestamo["cuota_mensual"],
        "cuota_anual": prestamo["cuota_anual"], "monto_prestamo": prestamo["monto"],
        "mezcla": equilibrio["mezcla"]["valor"], "utilidad": equilibrio["utilidad"]["valor"],
        "margen_baguette": de(("margen de contribucion", "baguette")),
        "margen_cupcake": de(("margen de contribucion", "cupcake")),
        "margen_ponderado": de(("margen de contribucion ponderado",)),
        "equilibrio_base": de(("punto de equilibrio base",)),
        "equilibrio_ajustado": de(("punto de equilibrio ajustado",)),
        "van": tir["van"], "tir": tir["tir"], "recuperacion": tir["recuperacion"],
    }


def actualizar_supuesto(campo, valor, origen=ORIGEN_ROBOT):
    """Cambia UNO de los cuatro supuestos financieros que el mapa marca como manuales.

    campo: 'tasa', 'plazo', 'mezcla' o 'utilidad' (ver SUPUESTOS).
    valor: como lo digita la persona: el porcentaje (12.51, 50, 20) o los meses (60).
    Devuelve ademas 'indicadores_antes' e 'indicadores' (VAN, TIR, punto de equilibrio...) y 'avisos'.
    """
    clave, definicion = _definicion_supuesto(campo)
    nuevo, _ = normalizar_supuesto(clave, valor)

    def _editar(wb, wb_valores):
        celda = _celda_supuesto(wb[definicion["hoja"]], clave)
        anterior = celda.value
        if _es_formula(anterior):
            raise FinanzasInvalido(f"'{definicion['hoja']}'!{celda.coordinate} es una formula: el robot no la sobrescribe.")
        if _numero(anterior) is None:
            raise FinanzasInvalido(f"'{definicion['hoja']}'!{celda.coordinate} trae {anterior!r}, que no es un numero: "
                                   "el robot no lo sobrescribe.")
        base = {"hoja": definicion["hoja"], "celda": celda.coordinate, "campo": clave, "titulo": definicion["titulo"],
                "antes": anterior, "despues": nuevo, "antes_texto": _como_se_ve(clave, anterior),
                "despues_texto": _como_se_ve(clave, nuevo), "avisos": []}
        if round(float(anterior), 4) == round(float(nuevo), 4):
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", **base}
        antes = leer_indicadores(wb_valores)
        celda.value = nuevo
        evento = registrar_evento(
            wb, modulo=MODULO, tipo_evento=definicion["evento"], entidad=definicion["hoja"], id_entidad=celda.coordinate,
            accion="UPDATE", estado="OK", origen=origen or ORIGEN_ROBOT, valor_referencia=nuevo,
            detalle=(f"'{definicion['hoja']}'!{celda.coordinate} {definicion['etiqueta']}: {anterior!r} -> {nuevo!r} "
                     f"({base['antes_texto']} -> {base['despues_texto']})"))
        return {"hubo_cambios": True, "estado": "OK", "evento": evento, "indicadores_antes": antes, **base}

    resultado = editar_excel(_editar, descripcion=f"{definicion['etiqueta']}: {nuevo}")
    if resultado.get("hubo_cambios"):
        # Segunda lectura: lo que calculo el Excel con el dato nuevo (VAN, TIR, punto de equilibrio).
        despues = leer_indicadores(leer_excel())
        resultado["indicadores"] = despues
        if clave == "plazo" and nuevo != PLAZO_DEL_MODELO:
            resultado["avisos"].append(
                f"AVISO: el flujo de caja del proyecto resta la cuota anual en los cinco años (fila 21), sea cual sea "
                f"el plazo. Con {nuevo} meses la proyeccion de cinco años no coincide con los pagos del prestamo; el "
                f"modelo esta hecho para {PLAZO_DEL_MODELO} meses. Se guardo igual.")
        margenes = [despues[m] for m in ("margen_baguette", "margen_cupcake", "margen_ponderado")]
        if definicion["hoja"] == HOJA_EQUILIBRIO and any(_numero(m) is not None and m < 0 for m in margenes):
            resultado["avisos"].append(
                "AVISO: hay un margen de contribucion negativo (el costo unitario de la hornada cargada supera al "
                "precio): con esa hornada la operacion deja perdida y el punto de equilibrio no es alcanzable. "
                "Se guardo igual.")
    return resultado


def leer_flujo_proyecto(wb):
    """La hoja 'flujo caja proyecto' tal como la calculo el Excel (solo LECTURA; wb = libro de VALORES).

    Devuelve {'columnas': [8 titulos], 'filas': [...], 'prestamo': {...}, 'capital': [...], 'capital_total'}.
    """
    ws = wb[HOJA_PROYECTO]
    encabezado = [f for f in range(1, ws.max_row + 1) if _clave(ws.cell(f, 1).value) == "periodo"]
    final = _filas_que_dicen(ws, ("flujo de caja",), empieza=True)
    final = [f for f in final if encabezado and f > encabezado[0]]
    if len(encabezado) != 1 or not final:
        raise KeyError(f"'{HOJA_PROYECTO}': no encuentro la fila 'PERIODO' o la fila 'FLUJO DE CAJA': el mapa y el "
                       "Excel no coinciden.")
    fila_titulos, fila_variedad, fila_final = encabezado[0], encabezado[0] + 1, final[0]
    letras = "BCDEFGHI"
    columnas = []
    for letra in letras:
        variedad = " ".join(str(ws[f"{letra}{fila_variedad}"].value or "").split())
        titulo = " ".join(str(ws[f"{letra}{fila_titulos}"].value or "").split())
        columnas.append(variedad or titulo)
    filas = []
    for fila in range(fila_variedad + 1, fila_final + 1):
        etiqueta = " ".join(str(ws.cell(fila, 1).value or "").split())
        if not etiqueta:
            continue
        valores = [_numero(ws[f"{letra}{fila}"].value) for letra in letras]
        filas.append({"fila": fila, "etiqueta": etiqueta, "valores": valores,
                      "clase": "titulo" if all(v is None for v in valores) else "formula"})

    celda_tasa, celda_plazo = _celda_supuesto(ws, "tasa"), _celda_supuesto(ws, "plazo")
    fila_cuota = _filas_que_dicen(ws, ("amortizacion", "mensual"), desde=fila_final)
    fila_monto = _filas_que_dicen(ws, ("monto", "prestamo"), desde=fila_final)
    if len(fila_cuota) != 1 or len(fila_monto) != 1:
        raise KeyError(f"'{HOJA_PROYECTO}': no encuentro la fila de la cuota mensual o la del monto del prestamo.")
    prestamo = {"tasa": _numero(celda_tasa.value), "celda_tasa": celda_tasa.coordinate,
                "plazo": _numero(celda_plazo.value), "celda_plazo": celda_plazo.coordinate,
                "cuota_mensual": _numero(ws[f"B{fila_cuota[0]}"].value), "celda_cuota_mensual": f"B{fila_cuota[0]}",
                "cuota_anual": _numero(ws[f"C{fila_cuota[0]}"].value), "celda_cuota_anual": f"C{fila_cuota[0]}",
                "monto": _numero(ws[f"B{fila_monto[0]}"].value), "celda_monto": f"B{fila_monto[0]}"}

    capital, capital_total, celda_total = [], None, None
    inicio = _filas_que_dicen(ws, ("capital de trabajo",), desde=fila_monto[0] + 1, empieza=True)
    if inicio:
        for fila in range(inicio[0] + 1, ws.max_row + 1):
            etiqueta = " ".join(str(ws.cell(fila, 1).value or "").split())
            if not etiqueta:
                continue
            if _clave(etiqueta) == "total":
                capital_total, celda_total = _numero(ws[f"B{fila}"].value), f"B{fila}"
                break
            capital.append({"fila": fila, "etiqueta": etiqueta, "valor": _numero(ws[f"B{fila}"].value)})
    return {"columnas": columnas, "filas": filas, "prestamo": prestamo, "capital": capital,
            "capital_total": capital_total, "celda_capital_total": celda_total, "plazo_del_modelo": PLAZO_DEL_MODELO}


def leer_punto_equilibrio(wb):
    """La hoja 'Punto de Equilibrio' tal como la calculo el Excel (solo LECTURA; wb = libro de VALORES).

    Devuelve {'filas': [...], 'mezcla': fila, 'utilidad': fila, 'nota'}. Cada fila:
    {'fila', 'celda', 'etiqueta', 'valor', 'clase': 'manual' | 'formula', 'formato': 'porcentaje' | 'unidades' | 'colones'}.
    """
    ws = wb[HOJA_EQUILIBRIO]
    manuales = {_celda_supuesto(ws, campo).row: campo for campo in ("mezcla", "utilidad")}
    filas, nota = [], ""
    for fila in range(2, ws.max_row + 1):
        etiqueta = " ".join(str(ws.cell(fila, 1).value or "").split())
        if not etiqueta:
            continue
        valor = _numero(ws.cell(fila, 2).value)
        if valor is None:
            if etiqueta.casefold().startswith("nota"):
                nota = etiqueta
            continue
        texto = _sin_acentos(etiqueta)
        if "(%)" in texto or texto.startswith("%"):
            formato = "porcentaje"
        elif "unidades" in texto:
            formato = "unidades"
        else:
            formato = "colones"
        filas.append({"fila": fila, "celda": f"B{fila}", "etiqueta": etiqueta, "valor": valor, "formato": formato,
                      "clase": "manual" if fila in manuales else "formula", "campo": manuales.get(fila)})
    return {"filas": filas, "nota": nota,
            "mezcla": next(f for f in filas if f["campo"] == "mezcla"),
            "utilidad": next(f for f in filas if f["campo"] == "utilidad")}


def leer_tir_var(wb):
    """La hoja 'TIR VAR' tal como la calculo el Excel (solo LECTURA): VAN, TIR, recuperacion y los flujos."""
    ws = wb[HOJA_TIR]

    def junto_a(texto, columna_etiqueta, columna_valor):
        for fila in range(1, ws.max_row + 1):
            etiqueta = _sin_acentos(ws.cell(fila, columna_etiqueta).value)
            if etiqueta == texto or (texto.endswith("*") and etiqueta.startswith(texto[:-1])):
                return _numero(ws.cell(fila, columna_valor).value), ws.cell(fila, columna_valor).coordinate
        raise KeyError(f"'{HOJA_TIR}': no encuentro la etiqueta '{texto}': el mapa y el Excel no coinciden.")

    van, celda_van = junto_a("van", 6, 7)
    tir, celda_tir = junto_a("tir", 6, 7)
    recuperacion, celda_recuperacion = junto_a("periodo de recuperacion*", 6, 7)
    inversion, celda_inversion = junto_a("io", 2, 3)
    tasa, celda_tasa = junto_a("i", 2, 3)
    flujos = []
    for fila in range(1, ws.max_row + 1):
        etiqueta = " ".join(str(ws.cell(fila, 2).value or "").split())
        if re.fullmatch(r"f\d+\s+\d{4}", etiqueta.casefold()):
            flujos.append({"fila": fila, "etiqueta": etiqueta, "año": etiqueta.split()[-1],
                           "flujo": _numero(ws.cell(fila, 3).value)})
    # Flujo acumulado (columnas J = periodo, I = flujo, L = acumulado): empieza en la inversion inicial (negativa).
    acumulado = []
    for fila in range(1, ws.max_row + 1):
        periodo = _numero(ws.cell(fila, 10).value)
        if periodo is not None and _numero(ws.cell(fila, 12).value) is not None:
            acumulado.append({"periodo": int(periodo), "flujo": _numero(ws.cell(fila, 9).value),
                              "acumulado": _numero(ws.cell(fila, 12).value)})
    return {"van": van, "celda_van": celda_van, "tir": tir, "celda_tir": celda_tir,
            "recuperacion": recuperacion, "celda_recuperacion": celda_recuperacion,
            "inversion": inversion, "celda_inversion": celda_inversion, "tasa": tasa, "celda_tasa": celda_tasa,
            "flujos": flujos, "acumulado": acumulado}


def leer_proyeccion_ventas(wb):
    """La hoja 'proyeccion ventas' (solo LECTURA): un renglon por concepto con Baguette y CupCake."""
    ws = wb[HOJA_PROYECCION]
    filas = []
    for fila in range(1, ws.max_row + 1):
        etiqueta = " ".join(str(ws.cell(fila, 2).value or "").split())
        baguette, cupcake = _numero(ws.cell(fila, 3).value), _numero(ws.cell(fila, 4).value)
        if etiqueta and baguette is not None:
            filas.append({"fila": fila, "etiqueta": etiqueta, "baguette": baguette, "cupcake": cupcake,
                          "dinero": "colones" in etiqueta.casefold() or "utilidad" in etiqueta.casefold()})
    extras = []
    for fila in range(1, ws.max_row + 1):                     # 'Precio de venta' y 'Costo unitario': el valor esta debajo
        etiqueta = " ".join(str(ws.cell(fila, 5).value or "").split())
        if etiqueta and _numero(ws.cell(fila + 1, 5).value) is not None and _numero(ws.cell(fila, 5).value) is None:
            extras.append({"fila": fila + 1, "etiqueta": etiqueta, "baguette": _numero(ws.cell(fila + 1, 5).value),
                           "cupcake": _numero(ws.cell(fila + 1, 6).value), "dinero": True})
    if not filas:
        raise KeyError(f"'{HOJA_PROYECCION}': no encuentro los renglones de la proyeccion: el mapa y el Excel no coinciden.")
    return {"filas": extras + filas, "utilidad_anual": _numero(ws["A1"].value), "celda_utilidad_anual": "A1"}


def leer_proyeccion_anual(wb):
    """La hoja 'proyeccion anual ventas' (solo LECTURA): unidades y ventas de cada mes y el total del año."""
    ws = wb[HOJA_PROYECCION_ANUAL]
    encabezado = [f for f in range(1, ws.max_row + 1) if _clave(ws.cell(f, 1).value) == "variables"]
    if len(encabezado) != 1:
        raise KeyError(f"'{HOJA_PROYECCION_ANUAL}': no encuentro la fila 'Variables': el mapa y el Excel no coinciden.")
    letras = "CDEFGHIJKLMN"
    meses = [" ".join(str(ws[f"{letra}{encabezado[0]}"].value or "").split()).capitalize() for letra in letras]
    filas, total, celda_total = [], None, None
    for fila in range(encabezado[0] + 1, ws.max_row + 1):
        etiqueta = " ".join(str(ws.cell(fila, 1).value or "").split())
        if not etiqueta:
            continue
        valores = [_numero(ws[f"{letra}{fila}"].value) for letra in letras]
        if valores[0] is not None and all(v is None for v in valores[1:]):
            total, celda_total, etiqueta_total = valores[0], f"C{fila}", etiqueta
            continue
        filas.append({"fila": fila, "etiqueta": etiqueta, "meses": valores, "dinero": "unidades" not in etiqueta.casefold()})
    return {"meses": meses, "filas": filas, "total": total, "celda_total": celda_total,
            "etiqueta_total": etiqueta_total if total is not None else ""}


# ----------------------------------------------------------------------------
# Uso manual desde la Terminal
# ----------------------------------------------------------------------------
def _main(argumentos):
    if argumentos == ["--ver"]:
        wb = leer_excel()
        diario = leer_flujo_diario(wb)
        print(f"'{HOJA_DIARIO}': primer dia {diario['primera']}, ultimo calculado {diario['ultima_calculada']}, "
              f"saldo de caja ({diario['celda_saldo']}) {diario['saldo']}")
        for dia in diario["dias"][:5]:
            print(f"   {dia['fecha']}  fila {dia['fila']:>3}  caja chica ({dia['celda_caja']}): {dia['caja_chica']}")
        anual = leer_flujo_anual(wb)
        salidas = anual["salidas"]
        print(f"'{HOJA_ANUAL}': fila {salidas['fila']} Salidas de efectivo: "
              + ", ".join(f"{m[:3]} {v if v is not None else '-'}" for m, v in zip(anual["meses"], salidas["meses"])))
        for nombre, valor in leer_indicadores(wb).items():
            print(f"   {nombre:<20} {valor}")
        return 0
    try:
        if len(argumentos) == 4 and argumentos[0] == "caja-chica":
            print(registrar_caja_chica(argumentos[1], argumentos[2], argumentos[3], origen="Terminal"))
        elif len(argumentos) == 3 and argumentos[0] == "salidas":
            print(registrar_salida_efectivo(argumentos[1], argumentos[2], origen="Terminal"))
        elif len(argumentos) == 3 and argumentos[0] == "supuesto":
            resultado = actualizar_supuesto(argumentos[1], argumentos[2], origen="Terminal")
            print({k: v for k, v in resultado.items() if not k.startswith("indicadores")})
            for aviso in resultado.get("avisos", []):
                print(aviso)
        else:
            print("Uso: py -m microservicios.grupo5_financiero.finanzas [--local] --ver\n"
                  "     py -m microservicios.grupo5_financiero.finanzas [--local] caja-chica <fecha> <monto> \"<motivo>\"\n"
                  "     py -m microservicios.grupo5_financiero.finanzas [--local] salidas <mes> <monto>\n"
                  f"     py -m microservicios.grupo5_financiero.finanzas [--local] supuesto <{' | '.join(SUPUESTOS)}> <valor>")
            return 2
    except FinanzasInvalido as error:
        print(f"RECHAZADO: {error}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(_main([a for a in sys.argv[1:] if a != "--local"]))
