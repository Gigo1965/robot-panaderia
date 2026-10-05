"""Microservicio: NODO INTEGRAL automatico (Grupo 3A).

Regla #4 del libro de mapas: NODO INTEGRAL es el UNICO proceso que consulta las
APIs externas. El SaaS y el bot solo LEEN lo que queda escrito aqui.

Que escribe (columna C, ubicando cada fila por su etiqueta de la columna B):
    Clima Actual            -> 'Lluvia' / 'Despejado'           (cada 30 min)
    Temperatura clima       -> numero en grados C               (cada 30 min)
    Tipo de Cambio          -> venta de referencia del BCCR      (cada 30 min)
    Calendario Feriados     -> nombre del feriado o 'Ninguno'    (1 vez al dia)
    Días de Pago            -> 'SI' / 'NO'                       (1 vez al dia)
    Feriado Activo Hoy      -> 'SI' / 'NO'                       (1 vez al dia)
    Día de Pago Activo Hoy  -> 'SI' / 'NO'                       (1 vez al dia)

Al recalcular, 'Tabla Combos Dinamicos' activa sola los combos segun estos valores.

Si una fuente falla: NO se borra el valor anterior; se anota el error en
RegistroEventos y las demas fuentes se actualizan igual.
Si nada cambio: no se sube el archivo (evita gastar versiones del historial de Drive).
"""
from conexion_drive import editar_excel
from conexion_drive.tablas import registrar_evento
from conexion_drive.tiempo import ahora_local

from . import fuentes

HOJA = "NODO INTEGRAL"
ETIQUETAS = {
    "clima": "Clima Actual",
    "temperatura": "Temperatura clima",
    "tipo_cambio": "Tipo de Cambio",
    "feriado_nombre": "Calendario Feriados",
    "dia_pago_lista": "Días de Pago",
    "feriado_hoy": "Feriado Activo Hoy",
    "dia_pago_hoy": "Día de Pago Activo Hoy",
}


def _celdas(ws):
    """{'clima': 'C5', ...} buscando cada etiqueta en la columna B."""
    encontradas = {}
    for fila in range(1, ws.max_row + 1):
        texto = str(ws.cell(fila, 2).value or "").strip().lower()
        for clave, etiqueta in ETIQUETAS.items():
            if texto == etiqueta.lower():
                encontradas[clave] = f"C{fila}"
    faltan = set(ETIQUETAS) - set(encontradas)
    if faltan:
        raise KeyError(f"No encontre en NODO INTEGRAL las filas: {sorted(ETIQUETAS[f] for f in faltan)}")
    return encontradas


def consultar_fuentes(clima=True, tipo_cambio=True, calendario=True, fecha=None):
    """Consulta internet y devuelve (valores, errores). No toca el Excel."""
    valores, errores = {}, {}
    if clima:
        try:
            dato = fuentes.obtener_clima()
            valores["clima"], valores["temperatura"] = dato["estado"], dato["temperatura"]
        except fuentes.FuenteNoDisponible as error:
            errores["clima"] = str(error)
    if tipo_cambio:
        try:
            dato = fuentes.obtener_tipo_cambio()
            valores["tipo_cambio"] = dato["venta"]
        except fuentes.FuenteNoDisponible as error:
            errores["tipo_cambio"] = str(error)
    if calendario:
        fecha = fecha or ahora_local().date()
        try:
            nombre = fuentes.nombre_feriado(fecha)
            valores["feriado_nombre"] = nombre or "Ninguno"
            valores["feriado_hoy"] = "SI" if nombre else "NO"
        except fuentes.FuenteNoDisponible as error:
            errores["feriado"] = str(error)
        pago = "SI" if fuentes.es_dia_de_pago(fecha) else "NO"
        valores["dia_pago_lista"], valores["dia_pago_hoy"] = pago, pago
    return valores, errores


def actualizar_nodo(clima=True, tipo_cambio=True, calendario=True, fecha=None, valores=None, errores=None):
    """Consulta las fuentes y escribe NODO INTEGRAL.

    'valores' y 'errores' permiten pasar datos ya consultados (lo usan las pruebas).
    """
    if valores is None:
        valores, errores = consultar_fuentes(clima, tipo_cambio, calendario, fecha)
    errores = errores or {}

    def _editar(wb, wb_valores):
        ws, wv = wb[HOJA], wb_valores[HOJA]
        celdas = _celdas(ws)
        cambios = {}
        for clave, valor in valores.items():
            celda = celdas[clave]
            if wv[celda].value != valor:
                cambios[clave] = (wv[celda].value, valor)
                ws[celda] = valor
        if errores:
            registrar_evento(wb, modulo="NODO INTEGRAL", tipo_evento="FUENTE_FALLIDA", entidad=HOJA,
                             id_entidad=", ".join(errores), accion="Se conservo el valor anterior",
                             estado="ERROR", detalle=" | ".join(f"{k}: {v}" for k, v in errores.items())[:500])
        if not cambios and not errores:
            return {"hubo_cambios": False, "cambios": {}, "errores": {}}
        return {"hubo_cambios": True, "cambios": cambios, "errores": errores}

    return editar_excel(_editar, descripcion="Actualizar NODO INTEGRAL")


def combos_activos():
    """Lee (sin escribir) que combos quedaron ACTIVOS despues de la ultima actualizacion."""
    from conexion_drive import leer_excel
    from microservicios.grupo1_captura_operativa.catalogo import Catalogo
    catalogo = Catalogo(leer_excel())
    return [c["codigo"] + " " + c["nombre"] for c in catalogo.combos if c["activo"]], catalogo.reglas


if __name__ == "__main__":
    # Uso:  py -m microservicios.grupo3_integraciones.nodo_integral            <- todo
    #       py -m microservicios.grupo3_integraciones.nodo_integral --rapido   <- solo clima y tipo de cambio
    import logging
    import sys
    logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")
    rapido = "--rapido" in sys.argv
    resultado = actualizar_nodo(calendario=not rapido)
    for clave, (antes, despues) in resultado["cambios"].items():
        print(f"  {ETIQUETAS[clave]}: {antes!r} -> {despues!r}")
    for clave, error in resultado["errores"].items():
        print(f"  ERROR {clave}: {error}")
    if not resultado["hubo_cambios"]:
        print("  Sin cambios: no se subio nada.")
    activos, reglas = combos_activos()
    print("Combos activos ahora:", activos or "ninguno",
          "| descuento general hoy:", "SI" if reglas["descuento_general"] else "NO")
    # || 4-oct-2026 (#84): con el Excel ya recalculado, el robot escribe 'Activacion de Promocion' (C16).
    from .activacion_promocion import actualizar_activacion_promocion
    promo = actualizar_activacion_promocion()
    print(f"Activacion de Promocion ({promo['celda']}): {promo['valor']} ({promo['estado']}) | {promo['motivo']}")
