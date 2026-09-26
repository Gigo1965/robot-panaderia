"""Migracion m006 - Cafe a la venta por paquete de 250 g (Decisiones Pendientes #25).

Decisiones del usuario (27-sep-2026):
  - El cafe se vende en despacho por PAQUETE de 250 g. Se compra en caja de 10
    paquetes, pero el inventario va en la UNIDAD DE VENTA (1 caja = 10 paquetes).
  - El cafe es COMPRA EXTERNA: se revende tal cual. Comprarlo en bulto no lo hace
    materia prima. Que saliera "Materia Prima" en Ingreso de Inventario L era un hueco.
  - La fila del cafe en TABLA CRUCE MAESTRA MP (fila 19) SE CONSERVA.
  - Mismo nombre tecnico que Control de Inventario A21 (un solo nombre, Regla #26).

Cambios (bloques):
  A) 'TABLA CRUCE MAESTRA' fila 15 (reserva vacia dentro de TablaProductos C6:I15):
     C nombre, D = MP!D19/10 (sigue al precio de la API de Solera), E 'Cf01',
     F 'Café', G = costo x 1.05 (igual que los demas de terceros), H, I.
  B) 'Ventas despacho': L13 = 'TABLA CRUCE MAESTRA'!C15 y lista desplegable
     E5:E27 de $L$2:$L$12 a $L$2:$L$13.
  C) 'Ingreso de Inventario' D19: 1 -> 10 (la caja de prueba en paquetes; el total
     sigue en 20,755 porque E19 pasa a tomar el costo por paquete).
  D) 'Ingreso de Inventario' L4:L37: primero pregunta si el producto se VENDE
     (TABLA CRUCE MAESTRA) -> "Compras Externas Netas"; si no, si esta en la MP ->
     "Materia Prima". Solo cambia el cafe (unico producto en ambas tablas).
  E) DiccionarioDatos: texto agregado en N56 (INV.ING.CATEGORIA) y fila nueva 305.
  F) Evento en RegistroEventos.

precios_proveedores.py NO toca D15: busca el codigo Ca01 solo en la MP (col F) y en
la TABLA CRUCE MAESTRA solo los 5 codigos de MaxiPali; Cf01 no esta en ninguna lista.

Uso (desde la carpeta ROBOT PANADERO):
    py -m migraciones.m006_cafe_venta --local    (copia datos_prueba/maestro_prueba.xlsx)
    py -m migraciones.m006_cafe_venta            (Excel real de Google Drive)
Si ya estaba aplicada, responde "Ya estaba aplicada" y no escribe nada.
"""
import os
import sys
from copy import copy

# --local ANTES de importar conexion_drive.config (igual que las demas migraciones)
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

NOMBRE_CAFE = "Café Naranjo puro 250gr x 10 Unidades"
CODIGO = "Cf01"
FECHA = "27-sep-2026"

TCM, MP = "TABLA CRUCE MAESTRA", "TABLA CRUCE MAESTRA MP"
VD, ING, CTRL, DIC = "Ventas despacho", "Ingreso de Inventario", "Control de Inventario", "DiccionarioDatos"

FILA_TCM = {
    "C15": NOMBRE_CAFE,
    "D15": "='TABLA CRUCE MAESTRA MP'!$D$19/10",
    "E15": CODIGO,
    "F15": "Café",
    "G15": "=TablaProductos[[#This Row],[Costo Unitario (Robot)]]*(1.05)",
    "H15": "Scraping Solera",
    "I15": "Tercero",
}
L_VIEJA = ('=IF(C{f}="","",IF(ISNUMBER(MATCH(C{f},TablaCruceMaestraMP[Materia Prima],0)),'
           '"Materia Prima","Compras Externas Netas"))')
L_NUEVA = ('=IF(C{f}="","",IF(ISNUMBER(MATCH(C{f},TablaProductos[Nombre tecnico completo],0)),'
           '"Compras Externas Netas",IF(ISNUMBER(MATCH(C{f},TablaCruceMaestraMP[Materia Prima],0)),'
           '"Materia Prima","Compras Externas Netas")))')
FILAS_L = range(4, 38)          # TablaIngresos A3:L37

TEXTO_N56 = (f" || V14 m006 ({FECHA}): la formula L pregunta PRIMERO si el producto se vende tal como se "
             "compra (esta en TABLA CRUCE MAESTRA) -> 'Compras Externas Netas'; si no, si esta en la MP -> "
             "'Materia Prima'. Corrige el cafe (se revende; comprarlo en caja de 10 no lo hace materia prima). "
             "Rango L4:L37.")
FILA_DIC = [
    "INV.TCM.CAFE_VENTA", "Inventario", TCM, "TablaProductos", "Fila 15 (C15:I15)",
    "Texto / numero CRC", "TABLA CRUCE MAESTRA MP D19 (API de catalogo Solera) / 10", "Formula",
    "Ventas despacho (lista L13) / Ingreso de Inventario E / catalogo del bot y SaaS", "No", "No", "—",
    "INV.ING.CATEGORIA",
    (f"CREADO m006 ({FECHA}, #25): el cafe se vende por paquete de 250 g (codigo Cf01). Costo por paquete "
     "= precio de la caja de 10 (MP D19) / 10; precio de venta = costo x 1.05. Inventario en unidad de venta: "
     "1 caja = 10 paquetes (Ingreso D19 de prueba 1 -> 10). Es COMPRA EXTERNA (se revende tal cual). La fila "
     "del cafe en TABLA CRUCE MAESTRA MP se conserva. Mismo nombre que Control de Inventario A21."),
]


class MigracionDetenida(Exception):
    """El Excel no esta como se esperaba: no se escribe nada."""


def _revisar(condicion, mensaje):
    if not condicion:
        raise MigracionDetenida(mensaje + " No se escribio nada.")


def ya_aplicada(wb):
    return str(wb[TCM]["E15"].value or "").strip() == CODIGO


def _validacion_despacho(ws):
    for dv in ws.data_validations.dataValidation:
        if "E5" in str(dv.sqref) and str(dv.formula1).replace(" ", "") == "$L$2:$L$12":
            return dv
    return None


def revisar_antes(wb):
    """Todas las revisiones ANTES de escribir la primera celda."""
    tcm, mp, ing, ctrl, vd = wb[TCM], wb[MP], wb[ING], wb[CTRL], wb[VD]
    _revisar(tcm.tables["TablaProductos"].ref == "C6:I15", "TablaProductos no es C6:I15.")
    for celda in FILA_TCM:
        _revisar(tcm[celda].value in (None, ""), f"TABLA CRUCE MAESTRA {celda} no esta vacia.")
    _revisar(mp["C19"].value == NOMBRE_CAFE and mp["F19"].value == "Ca01", "MP fila 19 no es el cafe Ca01.")
    _revisar(ctrl["A21"].value == NOMBRE_CAFE, "Control de Inventario A21 no es el cafe.")
    _revisar(ing["C19"].value == NOMBRE_CAFE and ing["D19"].value == 1,
             "Ingreso de Inventario fila 19 no es la compra de 1 caja de cafe.")
    _revisar(ing.tables["TablaIngresos"].ref == "A3:L37", "TablaIngresos no es A3:L37.")
    for f in FILAS_L:
        _revisar(str(ing[f"L{f}"].value).replace(" ", "") == L_VIEJA.format(f=f).replace(" ", ""),
                 f"Ingreso de Inventario L{f} no tiene la formula esperada.")
    _revisar(vd["L13"].value in (None, ""), "Ventas despacho L13 no esta vacia.")
    _revisar(_validacion_despacho(vd) is not None, "No encontre la lista desplegable E5 -> $L$2:$L$12.")
    _revisar(wb[DIC]["A56"].value == "INV.ING.CATEGORIA", "DiccionarioDatos A56 no es INV.ING.CATEGORIA.")
    _revisar(all(wb[DIC].cell(305, c).value in (None, "") for c in range(1, 15)), "DiccionarioDatos fila 305 no esta vacia.")


def _copiar_estilo(origen, destino):
    if origen.has_style:
        destino.font, destino.border, destino.fill = copy(origen.font), copy(origen.border), copy(origen.fill)
        destino.number_format, destino.alignment = origen.number_format, copy(origen.alignment)
        destino.protection = copy(origen.protection)


def aplicar(wb):
    """Devuelve el conteo de celdas por bloque."""
    conteo = {}
    tcm, ing, vd, dic = wb[TCM], wb[ING], wb[VD], wb[DIC]

    for celda, valor in FILA_TCM.items():                                   # A
        _copiar_estilo(tcm[celda.replace("15", "14")], tcm[celda])
        tcm[celda] = valor
    conteo["A"] = len(FILA_TCM)

    _copiar_estilo(vd["L12"], vd["L13"])                                    # B
    vd["L13"] = "='TABLA CRUCE MAESTRA'!C15"
    _validacion_despacho(vd).formula1 = "$L$2:$L$13"
    conteo["B"] = 2

    ing["D19"] = 10                                                         # C
    conteo["C"] = 1

    for f in FILAS_L:                                                       # D
        ing[f"L{f}"] = L_NUEVA.format(f=f)
    conteo["D"] = len(FILAS_L)

    dic["N56"] = str(dic["N56"].value or "") + TEXTO_N56                    # E
    for c, valor in enumerate(FILA_DIC, start=1):
        _copiar_estilo(dic.cell(304, c), dic.cell(305, c))
        dic.cell(305, c).value = valor
    conteo["E"] = 1 + len(FILA_DIC)
    return conteo


def main():
    from conexion_drive import config
    from conexion_drive.drive_excel import editar_excel
    from conexion_drive.tablas import registrar_evento

    destino = ("COPIA LOCAL " + str(config.EXCEL_LOCAL_PATH)) if config.ORIGEN_EXCEL == "local" \
        else "GOOGLE DRIVE (real)"
    print(f"=== m006_cafe_venta -> {destino} ===")

    def editar(wb, wb_valores):
        if ya_aplicada(wb):
            return {"hubo_cambios": False, "estado": "YA_APLICADA"}
        revisar_antes(wb)
        conteo = aplicar(wb)
        detalle = ", ".join(f"{b} {n}" for b, n in conteo.items())
        registrar_evento(
            wb, modulo="Migraciones", tipo_evento="MIGRACION", entidad="m006_cafe_venta",
            id_entidad="m006", accion="Cafe a la venta por paquete + clasificacion compra externa",
            estado="OK", detalle=f"Bloques: {detalle} (#25)", origen="Robot Python",
        )
        return {"hubo_cambios": True, "estado": "APLICADA", "conteo": conteo}

    try:
        resultado = editar_excel(editar, descripcion="Migracion m006_cafe_venta")
    except MigracionDetenida as error:
        print(f"DETENIDA: {error}")
        sys.exit(1)
    if resultado["estado"] == "YA_APLICADA":
        print("Ya estaba aplicada. No se escribio nada.")
        return
    for bloque, n in resultado["conteo"].items():
        print(f"  Bloque {bloque}: {n} celdas")
    print("LISTO.")


if __name__ == "__main__":
    main()
