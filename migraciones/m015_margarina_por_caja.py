"""Migracion m015 - La margarina Ma01 se vende POR CAJA; DiccionarioDatos al dia con C16 (5-oct-2026).

Dos pendientes pequenos que piden cambiar el Excel, juntos en UNA migracion (decision del usuario):

1) Decisiones Pendientes #36 (margarina Ma01). El producto se llama 'Margarina Numar caja de 4 Uds -
   500 g'; su costo en 'TABLA CRUCE MAESTRA' (columna 'Costo Unitario (Robot)') es el de la CAJA, que
   es lo que lee el robot en la API de catalogo; las compras y 'Control de Inventario' cuentan CAJAS.
   Lo unico que hablaba de barras era el precio de venta:  costo / 4 x 1.05  (el precio de UNA barra).
   El bot mostraba el nombre de la caja con el precio de la barra, y cada venta rebajaba 1 de un
   inventario que esta en cajas.
   DECISION DEL USUARIO (5-oct-2026): la margarina se vende por caja. El precio pasa a ser
   costo x 1.05,  la misma formula de los demas productos de reventa. No cambia el costo, el nombre,
   las compras ni el inventario. Lo ya vendido no se modifica (Regla #40): la unica venta guarda su
   precio como numero.

2) Decisiones Pendientes #86 (a). En 'DiccionarioDatos' la fila NODO.ACTIVACION_PROMOCION dice que
   'NODO INTEGRAL'!C16 la calcula el Excel con una regla interna. Desde el #84 la escribe el robot
   (activacion_promocion.py). Se agrega la aclaracion.

Que escribe:
  A) 'TABLA CRUCE MAESTRA', fila de Ma01, columna 'Precio': la formula deja de dividir entre 4.
  B) DiccionarioDatos, fila COST.PROD.COSTO_MARGARINA, columna 'Descripción': texto agregado.
  C) DiccionarioDatos, fila NODO.ACTIVACION_PROMOCION, columnas 'Origen', 'Escritura' y
     'Descripción': texto agregado.
  D) Evento MIGRACION en RegistroEventos.

No se borra nada: en el diccionario el texto viejo se conserva y lo nuevo va despues de '|| m015'.
Las filas se ubican por su CODIGO (Ma01) y por su ID del diccionario, no por numero fijo de fila.

Esperado en comparar_libros: 'TABLA CRUCE MAESTRA' 1 celda (la formula del precio de Ma01, que con
el costo de hoy pasa de 262.50 a 1,050.00); DiccionarioDatos 4 celdas; RegistroEventos 1 fila
(11 celdas); 16 celdas distintas en total; 0 valores calculados cambiados (ninguna otra formula lee
ese precio) salvo, si el archivo no se recalculaba desde otro dia, los 5 de la fila del dia de
'Flujo caja diario' por cada dia; 0 errores nuevos.

Revisiones ANTES de escribir (si una falla, se detiene y no escribe nada):
  - Ma01 aparece UNA sola vez en TablaProductos y su costo es un numero mayor que 0.
  - La formula del precio de Ma01 es exactamente la vieja (costo / 4 x 1.05).
  - Las dos filas del diccionario existen UNA sola vez y sus celdas son texto sin '|| m015'.
  - Ningun texto nuevo empieza con '='.

Uso (desde la carpeta ROBOT PANADERO):
    py -m migraciones.m015_margarina_por_caja --local   (copia datos_prueba/maestro_prueba.xlsx)
    py -m migraciones.m015_margarina_por_caja           (Excel MAESTRO de Google Drive)
    py -m migraciones.m015_margarina_por_caja --demo    (COPIA DE DEMOSTRACION de Google Drive)
Si ya estaba aplicada, responde "Ya estaba aplicada" y no escribe nada.

--demo usa la misma configuracion del SaaS (saas_demo/config.py), igual que la m014.
"""
import os
import sys

if "--local" in sys.argv and "--demo" in sys.argv:
    sys.exit("Use --local o --demo, no los dos.")
# --local ANTES de importar conexion_drive.config (igual que las demas migraciones)
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

FECHA = "5-oct-2026"
ID_MIGRACION = "m015_margarina_por_caja"
MARCA = "|| m015"
HOJA_TCM, DIC = "TABLA CRUCE MAESTRA", "DiccionarioDatos"
TABLA = "TablaProductos"
CODIGO = "Ma01"
COL_CODIGO, COL_COSTO, COL_PRECIO, COL_NOMBRE = "Codigo", "Costo Unitario (Robot)", "Precio", "Nombre tecnico completo"

FORMULA_VIEJA = "=TablaProductos[[#This Row],[Costo Unitario (Robot)]]/4*(1.05)"
FORMULA_NUEVA = "=TablaProductos[[#This Row],[Costo Unitario (Robot)]]*(1.05)"

ID_MARGARINA, ID_ACTIVACION = "COST.PROD.COSTO_MARGARINA", "NODO.ACTIVACION_PROMOCION"
# (ID de la fila del diccionario, encabezado de la columna, texto que se AGREGA al final)
NOTAS = [
    (ID_MARGARINA, "Descripción",
     f" {MARCA} ({FECHA}, #36 RESUELTO): la margarina se VENDE POR CAJA de 4 (decision del usuario). "
     "El precio de venta de Ma01 ya NO divide entre 4: pasa de costo / 4 x 1.05 a costo x 1.05, la misma "
     "formula de los demas productos de reventa. El nombre, el costo, las compras y el inventario ya "
     "estaban en cajas y no cambian. Lo ya vendido no se modifica (Regla #40)."),
    (ID_ACTIVACION, "Origen",
     f" {MARCA} ({FECHA}): Robot (activacion_promocion.py), #84."),
    (ID_ACTIVACION, "Escritura",
     f" {MARCA} ({FECHA}): Python. La escribe solo el robot; en el SaaS es de solo lectura (sin boton)."),
    (ID_ACTIVACION, "Descripción",
     f" {MARCA} ({FECHA}, #84 y #86 a): la celda NO tiene formula: la escribe el robot "
     "(activacion_promocion.py) con ON u OFF. ON = hoy hay una oferta: algun combo ACTIVO, o descuento "
     "general por dia de pago o feriado (Regla #8: esos dias no hay combos). OFF = ninguna de las dos. "
     "Se revisa en la tarea diaria, despues de actualizar NODO INTEGRAL, y al guardar la configuracion de "
     "un combo desde el SaaS. Si el valor no cambia, no se escribe nada."),
]


class MigracionDetenida(Exception):
    """El Excel no esta como se esperaba: no se escribe nada."""


def _revisar(condicion, mensaje):
    if not condicion:
        raise MigracionDetenida(mensaje + " No se escribio nada.")


def _sin_espacios(valor):
    return str(valor or "").replace(" ", "")


def _celdas_producto(wb):
    """Las celdas de Ma01 en TablaProductos: {'fila', 'nombre', 'costo', 'precio'} (celdas de openpyxl)."""
    from conexion_drive.tablas import columnas_de_tabla, filas_de_datos, obtener_tabla

    ws, tabla = obtener_tabla(wb, HOJA_TCM, TABLA)
    columnas = columnas_de_tabla(tabla)
    faltan = [c for c in (COL_CODIGO, COL_COSTO, COL_PRECIO, COL_NOMBRE) if c not in columnas]
    _revisar(not faltan, f"En {TABLA} faltan las columnas {faltan}.")
    filas = [f for f in filas_de_datos(ws, tabla)
             if str(ws[f"{columnas[COL_CODIGO]}{f}"].value or "").strip() == CODIGO]
    _revisar(len(filas) == 1, f"El codigo {CODIGO} aparece {len(filas)} veces en {TABLA} (debe aparecer una).")
    fila = filas[0]
    return {"fila": fila, "nombre": ws[f"{columnas[COL_NOMBRE]}{fila}"],
            "costo": ws[f"{columnas[COL_COSTO]}{fila}"], "precio": ws[f"{columnas[COL_PRECIO]}{fila}"]}


def _celdas_diccionario(wb):
    """Una celda de openpyxl por cada nota de NOTAS, en el mismo orden."""
    dic = wb[DIC]
    encabezado = {str(c.value).strip(): c.column for c in dic[4] if c.value is not None}
    _revisar("ID variable" in encabezado, "DiccionarioDatos fila 4 no es el encabezado.")
    filas_por_id = {}
    for celda in dic["A"]:
        if celda.row > 4 and celda.value:
            filas_por_id.setdefault(str(celda.value).strip(), []).append(celda.row)
    celdas = []
    for id_fila, columna, _ in NOTAS:
        _revisar(len(filas_por_id.get(id_fila, [])) == 1,
                 f"En DiccionarioDatos el ID {id_fila} aparece {len(filas_por_id.get(id_fila, []))} veces (debe aparecer una).")
        _revisar(columna in encabezado, f"DiccionarioDatos no tiene la columna '{columna}'.")
        celdas.append(dic.cell(filas_por_id[id_fila][0], encabezado[columna]))
    return celdas


def ya_aplicada(wb):
    precio = _celdas_producto(wb)["precio"]
    notas = [isinstance(c.value, str) and MARCA in c.value for c in _celdas_diccionario(wb)]
    formula_nueva = _sin_espacios(precio.value) == _sin_espacios(FORMULA_NUEVA)
    if formula_nueva and all(notas):
        return True
    _revisar(not formula_nueva and not any(notas),
             f"La migracion esta aplicada A MEDIAS ('{HOJA_TCM}'!{precio.coordinate} = {precio.value!r}; "
             f"notas del diccionario: {notas}). Hay que revisarlo antes de seguir.")
    return False


def revisar_antes(wb):
    producto = _celdas_producto(wb)
    costo, precio = producto["costo"], producto["precio"]
    _revisar(isinstance(costo.value, (int, float)) and not isinstance(costo.value, bool) and costo.value > 0,
             f"'{HOJA_TCM}'!{costo.coordinate} (costo de {CODIGO}) deberia ser un numero mayor que 0 y es {costo.value!r}.")
    _revisar(_sin_espacios(precio.value) == _sin_espacios(FORMULA_VIEJA),
             f"'{HOJA_TCM}'!{precio.coordinate} deberia ser {FORMULA_VIEJA} y es {precio.value!r}.")
    celdas = _celdas_diccionario(wb)
    for celda in celdas:
        _revisar(isinstance(celda.value, str) and celda.value.strip() and not celda.value.startswith("="),
                 f"DiccionarioDatos!{celda.coordinate} deberia ser un texto y es {celda.value!r}.")
    _revisar(not any(texto.strip().startswith("=") for _, _, texto in NOTAS), "Hay un texto que empieza con '='.")
    return {"producto": producto, "celdas_dic": celdas}


def aplicar(info):
    info["producto"]["precio"].value = FORMULA_NUEVA                      # A) el precio ya no divide entre 4
    for celda, (_, _, texto) in zip(info["celdas_dic"], NOTAS):            # B y C) texto agregado, nada se borra
        celda.value = celda.value.rstrip() + texto
    return {"A": 1, "B": 1, "C": len(NOTAS) - 1}


def verificar(wb):
    _revisar(_sin_espacios(_celdas_producto(wb)["precio"].value) == _sin_espacios(FORMULA_NUEVA),
             "Verificacion: la formula del precio no quedo como se esperaba.")
    _revisar(all(MARCA in c.value for c in _celdas_diccionario(wb)), "Verificacion: notas del diccionario.")


def main():
    demo = "--demo" in sys.argv
    palabra_copia = "DEMO"
    if demo:
        # La misma configuracion del SaaS: valida SAAS_EXCEL_FILE_ID (nunca el maestro) y deja al
        # robot apuntando a la copia. Debe importarse ANTES de conexion_drive.
        try:
            from saas_demo import config as config_saas
        except Exception as error:      # noqa: BLE001 (falta una variable SAAS_ o apunta al maestro)
            sys.exit(f"DETENIDA: la configuracion de la copia de demostracion no es valida: {error} "
                     "No se escribio nada.")
        if config_saas.ENSAYO:
            sys.exit("--demo no se usa en modo ensayo. No se escribio nada.")
        palabra_copia = config_saas.PALABRA_COPIA

    from conexion_drive import config
    from conexion_drive.drive_excel import editar_excel, obtener_metadatos
    from conexion_drive.tablas import registrar_evento

    if config.ORIGEN_EXCEL == "local":
        destino = "COPIA LOCAL " + str(config.EXCEL_LOCAL_PATH)
    else:
        nombre = str(obtener_metadatos().get("name") or "")
        es_copia = palabra_copia.lower() in nombre.lower()
        if demo and not es_copia:
            sys.exit(f"DETENIDA: el archivo de Drive se llama '{nombre}' y no contiene la palabra "
                     f"'{palabra_copia}': no parece la copia de demostracion. No se escribio nada.")
        if not demo and es_copia:
            sys.exit(f"DETENIDA: el archivo de Drive se llama '{nombre}': parece la copia de demostracion. "
                     "Para ella se usa --demo. No se escribio nada.")
        destino = f"GOOGLE DRIVE, {'COPIA DE DEMOSTRACION' if demo else 'EXCEL MAESTRO'}: {nombre}"
    print(f"=== {ID_MIGRACION} -> {destino} ===")

    def editar(wb, wb_valores):
        if ya_aplicada(wb):
            return {"hubo_cambios": False, "estado": "YA_APLICADA"}
        info = revisar_antes(wb)
        producto = info["producto"]
        costo = float(producto["costo"].value)
        precio_antes = wb_valores[HOJA_TCM][producto["precio"].coordinate].value
        precio_nuevo = round(costo * 1.05, 2)
        datos = {"celda_precio": producto["precio"].coordinate, "nombre": str(producto["nombre"].value),
                 "costo": costo, "precio_antes": precio_antes, "precio_nuevo": precio_nuevo,
                 "celdas_dic": [c.coordinate for c in info["celdas_dic"]]}
        conteo = aplicar(info)
        verificar(wb)
        antes = f"{precio_antes:,.2f}" if isinstance(precio_antes, (int, float)) else str(precio_antes)
        registrar_evento(
            wb, modulo="Migraciones", tipo_evento="MIGRACION", entidad=ID_MIGRACION, id_entidad="m015",
            accion="Margarina Ma01 se vende por caja; DiccionarioDatos al dia con C16",
            estado="OK",
            detalle=(f"A) '{HOJA_TCM}'!{datos['celda_precio']} precio de {CODIGO} ({datos['nombre']}): de costo / 4 x 1.05 "
                     f"a costo x 1.05 ({antes} -> {precio_nuevo:,.2f}; costo {costo:,.2f} sin cambio) | "
                     f"B) DiccionarioDatos {datos['celdas_dic'][0]}: nota agregada | "
                     f"C) DiccionarioDatos {', '.join(datos['celdas_dic'][1:])}: C16 la escribe el robot | "
                     "#36 resuelto, #86 (a), Regla #40"),
            origen="Robot Python",
        )
        return {"hubo_cambios": True, "estado": "APLICADA", "conteo": conteo, "datos": datos, "antes": antes}

    try:
        resultado = editar_excel(editar, descripcion=f"Migracion {ID_MIGRACION}")
    except MigracionDetenida as error:
        print(f"DETENIDA: {error}")
        sys.exit(1)
    if resultado["estado"] == "YA_APLICADA":
        print("Ya estaba aplicada. No se escribio nada.")
        return
    datos = resultado["datos"]
    print(f"  Bloque A: '{HOJA_TCM}'!{datos['celda_precio']} precio de {CODIGO}: {resultado['antes']} -> "
          f"{datos['precio_nuevo']:,.2f} (costo {datos['costo']:,.2f}, sin cambio)")
    print(f"  Bloque B: DiccionarioDatos {datos['celdas_dic'][0]} (nota de la margarina)")
    print(f"  Bloque C: DiccionarioDatos {', '.join(datos['celdas_dic'][1:])} (C16 la escribe el robot)")
    for bloque, n in resultado["conteo"].items():
        print(f"  Bloque {bloque}: {n} celdas")
    print("LISTO.")


if __name__ == "__main__":
    main()
