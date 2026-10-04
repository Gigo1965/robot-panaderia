"""Microservicio: descuento general de dia de pago y feriado (Grupo 1, Decisiones Pendientes #62).

Mapa de celdas 'Interface digital', celda B5: "Descuento Dias de Pago / Feriado (productos
individuales). Numerico (%). 5%. Manual (input directo del usuario)". Es el % que el robot
aplica a TODOS los productos sueltos (despacho y bot) cuando NODO INTEGRAL marca Feriado o
Dia de Pago activo. No aplica a los combos: el combo ya trae su precio especial (Regla #8).

Decisiones del usuario (3-oct-2026):
  * El descuento MAXIMO es 10 %. Fuera de 0 a 10 se rechaza y no se escribe nada.
  * El descuento es sobre el PRECIO DE VENTA del producto, los dias que hay descuento.
  * Lo cambia solo la administracion (el SaaS exige ese permiso).

Reglas del proyecto que se respetan:
  * Es CONFIGURACION (como costeo.py): el robot actualiza el valor (UPDATE) y deja un evento
    en RegistroEventos con el VALOR ANTERIOR -> VALOR NUEVO (no se pierde informacion).
  * La celda se ubica por su ETIQUETA ('Dias de Pago o feriado', columna A), igual que la lee
    catalogo.py; nunca se escribe sobre una formula.
  * Si el valor nuevo es igual al actual, no se sube nada (estado SIN_CAMBIO).

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo1_captura_operativa.descuento_general --local 5
"""
import os
import sys

# --local: igual que costeo.py y las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import editar_excel, leer_excel          # noqa: E402
from conexion_drive.tablas import registrar_evento            # noqa: E402

HOJA = "Interface digital"
ETIQUETA = "Días de Pago o feriado"        # columna A; el valor esta en la columna B (hoy B5)
MAXIMO_PORCENTAJE = 10                     # decision del usuario (3-oct-2026)
TIPO_EVENTO = "CONFIG_DESCUENTO_GENERAL"


class DescuentoInvalido(Exception):
    """Dato rechazado por una validacion: no se escribio nada."""


def normalizar_porcentaje(valor):
    """'5', '5%', '7,5' o 5 -> 5.0 (en PORCENTAJE: 5 significa 5 %). Fuera de 0 a 10: rechaza."""
    if isinstance(valor, bool):
        raise DescuentoInvalido(f"{valor!r} no es un porcentaje.")
    texto = str(valor if valor is not None else "").strip().replace("%", "").replace(",", ".").strip()
    try:
        porcentaje = round(float(texto), 2)
    except ValueError:
        raise DescuentoInvalido(f"{valor!r} no es un porcentaje. Escriba un numero, por ejemplo 5.") from None
    if not 0 <= porcentaje <= MAXIMO_PORCENTAJE:
        raise DescuentoInvalido(
            f"El descuento debe estar entre 0 % y {MAXIMO_PORCENTAJE} % (llego {porcentaje:g} %). "
            "No se escribio nada.")
    return porcentaje


def _celda(ws):
    for fila in range(1, ws.max_row + 1):
        texto = ws.cell(fila, 1).value
        if isinstance(texto, str) and texto.strip().lower() == ETIQUETA.lower():
            return ws.cell(fila, 2)
    raise DescuentoInvalido(f"No encontre la etiqueta '{ETIQUETA}' en la hoja '{HOJA}'. No se escribio nada.")


def leer_descuento_general(wb_valores=None):
    """Porcentaje actual (5.0 = 5 %), leido del Excel."""
    wb_valores = wb_valores or leer_excel()
    valor = _celda(wb_valores[HOJA]).value
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise DescuentoInvalido(f"'{HOJA}' no tiene un numero en el descuento general ({valor!r}).")
    return round(float(valor) * 100, 2)


def actualizar_descuento_general(porcentaje, origen="SaaS"):
    """Cambia el % de descuento de dia de pago y feriado. porcentaje: 5 significa 5 %."""
    porcentaje = normalizar_porcentaje(porcentaje)
    nuevo = round(porcentaje / 100, 4)             # la celda guarda la fraccion: 0.05 = 5 %

    def _editar(wb, wb_valores):
        celda = _celda(wb[HOJA])
        anterior = celda.value
        if isinstance(anterior, str) and anterior.startswith("="):
            raise DescuentoInvalido(f"'{HOJA}'!{celda.coordinate} es una formula: el robot no la sobrescribe.")
        if isinstance(anterior, (int, float)) and not isinstance(anterior, bool) and round(anterior, 4) == nuevo:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", "celda": celda.coordinate, "porcentaje": porcentaje}
        celda.value = nuevo
        antes = f"{anterior * 100:g} %" if isinstance(anterior, (int, float)) else repr(anterior)
        evento = registrar_evento(
            wb, modulo="Interface digital", tipo_evento=TIPO_EVENTO, entidad=HOJA, id_entidad=celda.coordinate,
            accion="UPDATE", estado="OK", origen=origen, valor_referencia=nuevo,
            detalle=f"'{HOJA}'!{celda.coordinate}: {anterior!r} -> {nuevo!r} ({antes} -> {porcentaje:g} %)",
        )
        return {"hubo_cambios": True, "estado": "OK", "celda": celda.coordinate, "evento": evento,
                "antes": anterior, "despues": nuevo, "porcentaje": porcentaje}

    return editar_excel(_editar, descripcion=f"Descuento general {porcentaje:g} %")


if __name__ == "__main__":
    argumentos = [a for a in sys.argv[1:] if a != "--local"]
    if len(argumentos) != 1:
        sys.exit("Uso: py -m microservicios.grupo1_captura_operativa.descuento_general [--local] <porcentaje>")
    try:
        print(actualizar_descuento_general(argumentos[0], origen="Terminal"))
    except DescuentoInvalido as error:
        sys.exit(f"RECHAZADO: {error}")
