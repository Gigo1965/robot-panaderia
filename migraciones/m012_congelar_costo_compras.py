"""Migracion m012 - Congelar el costo de las compras ya registradas (Decisiones Pendientes #32, Parte C).

Decisiones del usuario (26/27-sep-2026):
  - El precio de una compra viene SIEMPRE de la TABLA CRUCE MAESTRA / MP (nadie lo
    inventa ni lo digita) y, una vez registrada la compra, ese precio NO cambia.
  - Solo compras (columna E 'Precio unitario'). La produccion propia (I 'Precio
    Unitario2' y J) sigue como formula; hornadas.py no se toca.
  - Las compras NUEVAS ya se escriben con el costo congelado (ingreso_inventario.py,
    Partes A y B, probadas el 28-sep-2026). Esta migracion congela las VIEJAS.

Por que: E era una formula viva a las tablas maestras. Cuando precios_proveedores.py
cambie un precio (7:00 a. m.), TODAS las compras pasadas se recalcularian y el
flujo de caja diario historico se moveria solo (problema de auditoria).

Cambios (bloques):
  A) 'Ingreso de Inventario' E: en las 18 filas de COMPRA (C con producto:
     filas 4-20 y 37), la formula se reemplaza por su VALOR actual (numero).
     F (= D x E), I, J y L siguen como formula. Las 16 filas de hornada (C vacia,
     21-36) no se tocan.
  B) DiccionarioDatos fila 49 (INV.ING.PRECIO_UNITARIO): texto agregado en H49 y
     N49 (no se borra nada). Guarda el texto de la formula vieja SIN el '=' inicial.
  C) Evento en RegistroEventos.

Revisiones ANTES de escribir (si una falla, se detiene y no escribe nada):
  - TablaIngresos es A3:L37; hay exactamente 18 filas de compra y 16 sin compra.
  - Las 18 tienen la formula vieja exacta en E.
  - El valor calculado de cada una es un numero > 0 e IGUAL al costo que la TCM
    tiene hoy para ese producto (la misma busqueda que la formula, por nombre).
Esperado en comparar_libros: 18 formulas cambiadas y 0 valores cambiados.

Uso (desde la carpeta ROBOT PANADERO):
    py -m migraciones.m012_congelar_costo_compras --local   (copia datos_prueba/maestro_prueba.xlsx)
    py -m migraciones.m012_congelar_costo_compras           (Excel real de Google Drive)
Si ya estaba aplicada, responde "Ya estaba aplicada" y no escribe nada.
"""
import os
import sys

# --local ANTES de importar conexion_drive.config (igual que las demas migraciones)
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

FECHA = "28-sep-2026"
ING, TCM, MP, DIC = "Ingreso de Inventario", "TABLA CRUCE MAESTRA", "TABLA CRUCE MAESTRA MP", "DiccionarioDatos"
FILAS = range(4, 38)               # TablaIngresos A3:L37 (datos 4-37)
COMPRAS_ESPERADAS, SIN_COMPRA_ESPERADAS = 18, 16
MARCA = "V15 m012"

# Copia literal de la formula vieja (una migracion es una foto: no depende del codigo actual)
E_VIEJA = ('=IF(C{f}="","",IFERROR(INDEX(TablaProductos[Costo Unitario (Robot)], MATCH(C{f}, '
           'TablaProductos[Nombre tecnico completo], 0)), IFERROR(INDEX(TablaCruceMaestraMP[Costo Unitario '
           '(Scraping)], MATCH(C{f}, TablaCruceMaestraMP[Materia Prima], 0)), 0)))')

TEXTO_H49 = f" || {MARCA} ({FECHA}): Robot al registrar la compra (numero congelado, no formula)."
TEXTO_N49 = (f" || {MARCA} ({FECHA}, #32): el costo de cada COMPRA se CONGELA. El robot lo toma de la TCM al "
             "registrar la compra (misma busqueda de siempre: TablaProductos 'Costo Unitario (Robot)' y, si no, "
             "TablaCruceMaestraMP 'Costo Unitario (Scraping)', por nombre) y lo escribe como NUMERO en E. Asi un "
             "cambio de precio del proveedor no mueve las compras ya registradas. F = D x E sigue como formula; "
             "la produccion (I, J) sigue como formula. La m012 congelo con su valor las 18 compras existentes "
             "(filas 4-20 y 37). Formula anterior de E (sin el signo igual): ")


class MigracionDetenida(Exception):
    """El Excel no esta como se esperaba: no se escribe nada."""


def _revisar(condicion, mensaje):
    if not condicion:
        raise MigracionDetenida(mensaje + " No se escribio nada.")


def _limpio(texto):
    return str(texto or "").replace(" ", "")


def _es_numero(valor):
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


def ya_aplicada(wb):
    return MARCA in str(wb[DIC]["N49"].value or "")


def costos_tcm(wb_valores):
    """Costo actual por NOMBRE: primero TablaProductos, despues la MP (igual que la formula vieja)."""
    costos = {}
    for hoja, tabla, col_nombre, col_costo in ((TCM, "TablaProductos", "Nombre tecnico completo",
                                                "Costo Unitario (Robot)"),
                                               (MP, "TablaCruceMaestraMP", "Materia Prima",
                                                "Costo Unitario (Scraping)")):
        ws = wb_valores[hoja]
        ref = ws.tables[tabla].ref
        filas = list(ws[ref])
        encabezado = [str(c.value or "").strip() for c in filas[0]]
        i_nombre, i_costo = encabezado.index(col_nombre), encabezado.index(col_costo)
        for fila in filas[1:]:
            nombre = str(fila[i_nombre].value or "").strip()
            if nombre and nombre not in costos:
                costos[nombre] = fila[i_costo].value
    return costos


def revisar_antes(wb, wb_valores):
    """Todas las revisiones ANTES de escribir. Devuelve [(fila, producto, valor), ...] de las compras."""
    ing, ing_v = wb[ING], wb_valores[ING]
    _revisar(ing.tables["TablaIngresos"].ref == "A3:L37", "TablaIngresos no es A3:L37.")
    _revisar(str(ing["E3"].value).strip() == "Precio unitario", "Ingreso de Inventario E3 no es 'Precio unitario'.")
    costos = costos_tcm(wb_valores)
    compras, sin_compra = [], 0
    for f in FILAS:
        producto = str(ing[f"C{f}"].value or "").strip()
        if not producto:
            sin_compra += 1
            continue
        _revisar(_limpio(ing[f"E{f}"].value) == _limpio(E_VIEJA.format(f=f)),
                 f"Ingreso de Inventario E{f} ({producto}) no tiene la formula vieja esperada.")
        valor = ing_v[f"E{f}"].value
        _revisar(_es_numero(valor) and valor > 0, f"E{f} ({producto}) no tiene un valor calculado > 0: {valor!r}.")
        _revisar(valor == costos.get(producto),
                 f"E{f} ({producto}) = {valor!r} no coincide con el costo de la TCM ({costos.get(producto)!r}).")
        compras.append((f, producto, valor))
    _revisar(len(compras) == COMPRAS_ESPERADAS,
             f"Hay {len(compras)} filas de compra (se esperaban {COMPRAS_ESPERADAS}).")
    _revisar(sin_compra == SIN_COMPRA_ESPERADAS,
             f"Hay {sin_compra} filas sin compra (se esperaban {SIN_COMPRA_ESPERADAS}).")
    dic = wb[DIC]
    _revisar(dic["A49"].value == "INV.ING.PRECIO_UNITARIO", "DiccionarioDatos A49 no es INV.ING.PRECIO_UNITARIO.")
    return compras


def aplicar(wb, compras):
    """Devuelve el conteo de celdas por bloque."""
    ing, dic = wb[ING], wb[DIC]
    for f, _, valor in compras:                                             # A
        ing[f"E{f}"] = valor            # el formato de la celda (estilo) no cambia
    dic["H49"] = str(dic["H49"].value or "") + TEXTO_H49                    # B
    dic["N49"] = str(dic["N49"].value or "") + TEXTO_N49 + E_VIEJA.format(f=4).lstrip("=")
    return {"A": len(compras), "B": 2}


def main():
    from conexion_drive import config
    from conexion_drive.drive_excel import editar_excel
    from conexion_drive.tablas import registrar_evento

    destino = ("COPIA LOCAL " + str(config.EXCEL_LOCAL_PATH)) if config.ORIGEN_EXCEL == "local" \
        else "GOOGLE DRIVE (real)"
    print(f"=== m012_congelar_costo_compras -> {destino} ===")

    def editar(wb, wb_valores):
        if ya_aplicada(wb):
            return {"hubo_cambios": False, "estado": "YA_APLICADA"}
        compras = revisar_antes(wb, wb_valores)
        conteo = aplicar(wb, compras)
        total = sum(f_v for _, _, f_v in compras)
        registrar_evento(
            wb, modulo="Migraciones", tipo_evento="MIGRACION", entidad="m012_congelar_costo_compras",
            id_entidad="m012", accion="Costo de compras congelado (formula -> numero) en Ingreso de Inventario E",
            estado="OK",
            detalle=(f"Bloques: A {conteo['A']}, B {conteo['B']} (#32). Filas: "
                     + ", ".join(str(f) for f, _, _ in compras)),
            origen="Robot Python",
        )
        return {"hubo_cambios": True, "estado": "APLICADA", "conteo": conteo, "compras": compras,
                "suma_precios": total}

    try:
        resultado = editar_excel(editar, descripcion="Migracion m012_congelar_costo_compras")
    except MigracionDetenida as error:
        print(f"DETENIDA: {error}")
        sys.exit(1)
    if resultado["estado"] == "YA_APLICADA":
        print("Ya estaba aplicada. No se escribio nada.")
        return
    for f, producto, valor in resultado["compras"]:
        print(f"  E{f:<3} {valor:>12,.2f}  {producto[:45]}")
    for bloque, n in resultado["conteo"].items():
        print(f"  Bloque {bloque}: {n} celdas")
    print(f"  Filas congeladas: {len(resultado['compras'])} (esperadas {COMPRAS_ESPERADAS})")
    print("LISTO.")


if __name__ == "__main__":
    main()
