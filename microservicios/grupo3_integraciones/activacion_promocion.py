"""Microservicio: 'Activacion de Promocion' de NODO INTEGRAL, celda C16 (Grupo 3, Decisiones Pendientes #84).

Hallazgo (4-oct-2026): el mapa describe C16 como un dato "Bloqueado (Lectura) - regla interna"
con valor ON / OFF, pero en el Excel la celda tenia el texto de relleno "[ON / OFF]": ni el
robot ni una formula la escribian.

Decisiones del usuario (4-oct-2026):
  * C16 la escribe EL ROBOT (opcion C; no una formula del Excel ni el SaaS).
  * ON  = hay una oferta hoy: algun combo ACTIVO, o descuento general por dia de pago / feriado.
    OFF = ninguna de las dos.
  * El SaaS la muestra solo en la pantalla NODO INTEGRAL.
  * En la copia de demostracion la tarea diaria no corre: la administracion la actualiza con
    un boton del SaaS que llama a ESTA misma funcion.

Como decide: con el mismo catalogo que usan el bot y el despacho (catalogo.py), que ya aplica
la Regla #8: en feriado o dia de pago hay descuento general y NO hay combos, aunque la columna
'Detonador' de la tabla de combos siga diciendo ACTIVO (su formula solo mira la temperatura).
Por eso C16 no copia a C20: pregunta primero por el descuento general.

Cuando corre:
  * En el maestro: al final de la tarea diaria de NODO INTEGRAL (nodo_integral.py), DESPUES de
    que el clima y el calendario quedaron guardados y el Excel recalculo los combos.
  * En la copia de demostracion: con el boton de la pantalla NODO INTEGRAL (solo administracion).

Reglas del proyecto que se respetan:
  * Lee el Excel ya recalculado; escribe SOLO la celda C16 (la ubica por su etiqueta de la
    columna B) y deja un evento con el valor anterior -> nuevo y el motivo.
  * Si el valor no cambia, no se sube nada (estado SIN_CAMBIO): la mayoria de los dias solo lee.
  * Nunca escribe sobre una formula.

Uso manual desde la Terminal (carpeta ROBOT PANADERO); con --local trabaja sobre datos_prueba:
    py -m microservicios.grupo3_integraciones.activacion_promocion --local --ver
    py -m microservicios.grupo3_integraciones.activacion_promocion --local
"""
import os
import sys

# --local: igual que descuento_general.py y las migraciones. Debe ir ANTES de importar conexion_drive.
if __name__ == "__main__" and "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import editar_excel, leer_excel                                  # noqa: E402
from conexion_drive.tablas import registrar_evento                                   # noqa: E402
from microservicios.grupo1_captura_operativa.catalogo import Catalogo                # noqa: E402

HOJA = "NODO INTEGRAL"
ETIQUETA = "Activación de Promoción"
COL_ETIQUETA, COL_VALOR = 2, 3                 # columnas B y C
ON, OFF = "ON", "OFF"
COMBOS, DESCUENTO_GENERAL, NINGUNA = "COMBOS", "DESCUENTO_GENERAL", "NINGUNA"
TIPO_EVENTO = "ACTIVACION_PROMOCION"


class ActivacionInvalida(Exception):
    """No se pudo ubicar o escribir la celda: no se escribio nada."""


def _celda(ws):
    filas = [f for f in range(1, ws.max_row + 1)
             if isinstance(ws.cell(f, COL_ETIQUETA).value, str)
             and ws.cell(f, COL_ETIQUETA).value.strip().lower() == ETIQUETA.lower()]
    if len(filas) != 1:
        raise ActivacionInvalida(
            f"La etiqueta '{ETIQUETA}' aparece {len(filas)} veces en la columna B de '{HOJA}' "
            "(debe aparecer una sola vez). No se escribio nada.")
    return ws.cell(filas[0], COL_VALOR)


def evaluar(wb_valores):
    """Decide ON / OFF con el Excel ya recalculado. No escribe.

    Devuelve {'valor': 'ON' | 'OFF', 'tipo': COMBOS | DESCUENTO_GENERAL | NINGUNA,
              'combos': [{'codigo', 'nombre'}], 'porcentaje': 5.0, 'detonador': 'DIA_PAGO' | 'FERIADO' | None,
              'motivo': texto corto para el evento}
    """
    catalogo = Catalogo(wb_valores)
    reglas = catalogo.reglas
    if reglas["descuento_general"]:                       # Regla #8: ese dia no hay combos
        porcentaje = round(reglas["pct_general"] * 100, 2)
        dia = "feriado" if reglas["feriado"] else "dia de pago"
        return {"valor": ON, "tipo": DESCUENTO_GENERAL, "combos": [], "porcentaje": porcentaje,
                "detonador": reglas["detonador_general"],
                "motivo": f"descuento general de {porcentaje:g} % por {dia} (ese dia no hay combos, Regla #8)"}
    activos = [{"codigo": c["codigo"], "nombre": c["nombre"]} for c in catalogo.combos if c["activo"]]
    if activos:
        return {"valor": ON, "tipo": COMBOS, "combos": activos, "porcentaje": None, "detonador": None,
                "motivo": "combos activos: " + ", ".join(c["codigo"] for c in activos)}
    return {"valor": OFF, "tipo": NINGUNA, "combos": [], "porcentaje": None, "detonador": None,
            "motivo": "ningun combo activo y hoy no es feriado ni dia de pago"}


def leer_activacion(wb_valores=None):
    """Lo que dice C16 y lo que corresponde segun el Excel de este momento (solo lectura).

    Agrega a evaluar(): 'celda', 'valor_celda' (lo escrito hoy en C16), 'definido' (True si la
    celda ya dice ON u OFF) y 'al_dia' (True si lo escrito coincide con lo que corresponde).
    """
    wb_valores = wb_valores or leer_excel()
    celda = _celda(wb_valores[HOJA])
    escrito = celda.value.strip() if isinstance(celda.value, str) else celda.value
    estado = evaluar(wb_valores)
    estado.update(celda=celda.coordinate, valor_celda=escrito, definido=escrito in (ON, OFF),
                  al_dia=escrito == estado["valor"])
    return estado


def actualizar_activacion_promocion(origen="Robot Python"):
    """Escribe ON u OFF en 'Activacion de Promocion' (C16) si cambio. Devuelve lo decidido."""

    def _editar(wb, wb_valores):
        estado = evaluar(wb_valores)
        celda = _celda(wb[HOJA])
        anterior = celda.value
        if (isinstance(anterior, str) and anterior.startswith("=")) or type(anterior).__name__ == "ArrayFormula":
            raise ActivacionInvalida(f"'{HOJA}'!{celda.coordinate} es una formula: el robot no la sobrescribe.")
        estado.update(celda=celda.coordinate, antes=anterior)
        if isinstance(anterior, str) and anterior.strip() == estado["valor"]:
            return {"hubo_cambios": False, "estado": "SIN_CAMBIO", **estado}
        celda.value = estado["valor"]
        evento = registrar_evento(
            wb, modulo="NODO INTEGRAL", tipo_evento=TIPO_EVENTO, entidad=HOJA, id_entidad=celda.coordinate,
            accion="UPDATE", estado="OK", origen=origen, valor_referencia=estado["valor"],
            detalle=f"'{HOJA}'!{celda.coordinate} {ETIQUETA}: {anterior!r} -> {estado['valor']!r} | {estado['motivo']}",
        )
        return {"hubo_cambios": True, "estado": "OK", "evento": evento, **estado}

    return editar_excel(_editar, descripcion="Activacion de Promocion (NODO INTEGRAL)")


if __name__ == "__main__":
    argumentos = [a for a in sys.argv[1:] if a != "--local"]
    if argumentos == ["--ver"]:
        e = leer_activacion()
        print(f"  {e['celda']} dice {e['valor_celda']!r}; corresponde {e['valor']} ({e['motivo']})"
              f"{'' if e['al_dia'] else '  <- falta actualizar'}")
        sys.exit(0)
    if argumentos:
        sys.exit("Uso: py -m microservicios.grupo3_integraciones.activacion_promocion [--local] [--ver]")
    try:
        r = actualizar_activacion_promocion(origen="Terminal")
    except ActivacionInvalida as error:
        sys.exit(f"RECHAZADO: {error}")
    print(f"  {r['celda']}: {r['antes']!r} -> {r['valor']!r} ({r['estado']}) | {r['motivo']}")
