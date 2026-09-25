"""Migracion m003 - Grupo 3C: productos y precios de proveedores (MaxiPali / Solera).

Que hace (confirmado por el usuario el 24-sep-2026, opcion D1):
  A) "Leche Suli" -> "Leche Semidescremada Dos Pinos -1 L" en las 11 celdas donde aparece.
  B) 7 materias primas de Solera reemplazadas (el producto viejo ya no existe):
     nombre nuevo en TABLA CRUCE MAESTRA MP (col C), Ingreso de Inventario (col C)
     y Control de Inventario (col A). El nombre es la LLAVE de los SUMIF/MATCH,
     por eso se cambia en las tres hojas a la vez.
  C) Conversiones de TABLA CRUCE MAESTRA MP col E (tamano del empaque nuevo).
  D) Precios verificados el 24-sep-2026: MP col D y TABLA CRUCE MAESTRA col D.
  E) Notas en DiccionarioDatos (se AGREGA texto con "|| m003"; no se borra nada).

Seguridad (protocolo del proyecto):
  * Antes de escribir revisa TODAS las celdas. Si alguna no tiene el valor
    "antes" esperado (ni el "despues"), se detiene SIN cambiar nada.
  * Si una celda ya tiene el valor "despues", la salta (se puede correr 2 veces).
  * Despues de escribir, vuelve a leer cada celda y confirma el valor.
  * Ningun texto escrito empieza con "=" (solo las 6 formulas de la col E).

Uso (desde la carpeta ROBOT PANADERO):
    py -m migraciones.m003_proveedores_3c --local            (copia datos_prueba)
    py -m migraciones.m003_proveedores_3c                    (Excel real en Drive)
    py -m migraciones.m003_proveedores_3c --archivo "ruta.xlsx"   (un archivo suelto)
    Agregar --sin-precios para NO aplicar el bloque D.
"""
import os
import sys

# --local: igual que m001/m002. Debe ir ANTES de importar conexion_drive.config,
# porque config lee ORIGEN_EXCEL al importarse. (Corregido 24-sep-2026: la 1a
# version no tenia esto y con --local apuntaba a Drive.)
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

ID_MIGRACION = "m003_proveedores_3c"

LECHE_NUEVA = "Leche Semidescremada Dos Pinos -1 L"
LECHE_TCM = "Leche Suli Semidescremada 1.5% grasa  - 1 L"      # con guion
LECHE_VENTAS = "Leche Suli Semidescremada 1.5% grasa  1 L"     # sin guion (no se descontaba)

TCM = "TABLA CRUCE MAESTRA"
MP = "TABLA CRUCE MAESTRA MP"
ING = "Ingreso de Inventario"
CTRL = "Control de Inventario"
COMBOS = "Tabla Combos Dinamicos"
DIC = "DiccionarioDatos"

# ---------------------------------------------------------------- Bloque A
BLOQUE_A = [
    (TCM, "C10", LECHE_TCM, LECHE_NUEVA),
    (CTRL, "A7", LECHE_TCM, LECHE_NUEVA),
    (ING, "C4", LECHE_TCM, LECHE_NUEVA),
    (ING, "C7", LECHE_TCM, LECHE_NUEVA),
    ("Ventas despacho", "E22", LECHE_VENTAS, LECHE_NUEVA),
    ("Ventas despacho", "E42", LECHE_VENTAS, LECHE_NUEVA),
    ("Ventas despacho", "E44", LECHE_VENTAS, LECHE_NUEVA),
    ("Ventas despacho", "E147", LECHE_VENTAS, LECHE_NUEVA),
    (COMBOS, "D5", "3 Baguette + 1 Leche Suli", "3 Baguette + 1 " + LECHE_NUEVA),
    (COMBOS, "K5", "🔥 Combo 1 activo: 3 Baguette + 1 Leche Suli",
     "🔥 Combo 1 activo: 3 Baguette + 1 " + LECHE_NUEVA),
    (DIC, "N80",
     "Descripcion del contenido del combo, redactada para que el robot/bot de WhatsApp "
     "la use al armar el pedido (ej. '3 Baguette + 1 Leche Suli').",
     "Descripcion del contenido del combo, redactada para que el robot/bot de WhatsApp "
     "la use al armar el pedido (ej. '3 Baguette + 1 " + LECHE_NUEVA + "')."),
]

# ---------------------------------------------------------------- Bloque B
# (codigo, fila MP, fila Ingreso, fila Control, nombre viejo MP, nombre viejo Ing/Ctrl, nombre nuevo)
SOLERA = [
    ("Me01", 10, 10, 12, "MEJORANTE STAR PAN PLUS 3 KILOS", "MEJORANTE STAR PAN PLUS 5 KILOS",
     "MEJORANTE PANNIVITA PRO-PAN 1 KG HOMO-PLUS"),
    ("Sa01", 13, 13, 15, "SAL FINA DIAMANTE 25 KG COLOR VERDE. GRANO NORMAL", None,
     "SAL SARU COMUN ECONOMICA 25 KG SACO BRIN"),
    ("Le01", 15, 15, 17, "LEVADURA TEMPO SALADA 500 GR", None,
     "LEVADURA FRESCA GOLONDRINA 400 GR CIP"),
    ("Me02", 16, 16, 18, "MERMELADA GUAYABA F3 1 KG ANGEL MANGA", None,
     "MERMELADA GUAYABA HORNEO 5 KG CCR CAJA C2 (FRUTA TRITURADA)"),
    ("Bo01", 17, 17, 19, "BOLSAS PAPEL KRAFT 1 BAGUETTE 500 UNDS BIO", None,
     "BOLSAS PAPEL KRAFT 2 BAGUETTE 100 UNDS BIODEGRADABLE"),
    ("Bu01", 18, 18, 20, "BURBUJA #6 BIO 50 UND. 6″X6″ POSTRE, REPOSTERIA BAZAGO ", None,
     "BURBUJA FOAM #6 125 UNID POSTRE BL"),
    ("Ca01", 19, 19, 21, "CAFE NARANJO PURO 250 GR X 3 UNIDADES", "CAFE NARANJO PURO 250 GR X 10 UNIDADES",
     "Café Naranjo puro 250gr x 10 Unidades"),
]
BLOQUE_B = []
for _cod, f_mp, f_ing, f_ctrl, viejo_mp, viejo_inv, nuevo in SOLERA:
    viejo_inv = viejo_inv or viejo_mp
    BLOQUE_B += [(MP, f"C{f_mp}", viejo_mp, nuevo),
                 (ING, f"C{f_ing}", viejo_inv, nuevo),
                 (CTRL, f"A{f_ctrl}", viejo_inv, nuevo)]

# ---------------------------------------------------------------- Bloque C
BLOQUE_C = [
    (MP, "E10", "=D10/3000", "=D10/1000"),   # Mejorante 1 kg
    (MP, "E15", "=D15/500", "=D15/400"),     # Levadura 400 g
    (MP, "E16", "=D16/1000", "=D16/5000"),   # Mermelada 5 kg
    (MP, "E17", "=D17/500", "=D17/100"),     # Bolsas 100 und
    (MP, "E18", "=D18/50", "=D18/125"),      # Burbuja 125 und
    (MP, "E19", "=D19/750", "=D19/2500"),    # Cafe 250 g x 10
]

# ---------------------------------------------------------------- Bloque D
BLOQUE_D = [
    (MP, "D10", 5000, 1815), (MP, "D11", 25000, 17510), (MP, "D12", 24000, 16490),
    (MP, "D13", 25000, 3310), (MP, "D14", 25000, 28845), (MP, "D15", 1000, 735),
    (MP, "D16", 1000, 4140), (MP, "D17", 1000, 2400), (MP, "D18", 1000, 7275),
    (MP, "D19", 2500, 20755), (MP, "D20", 3600, 15390),
    (TCM, "D10", 800, 1050), (TCM, "D11", 100, 2300), (TCM, "D12", 750, 1000),
    (TCM, "D14", 100, 840),   # D13 (Coca Cola) ya vale 750: no cambia
]

# ---------------------------------------------------------------- Bloque E
# Se AGREGA al final del texto existente (nunca se borra informacion).
NOTA_N70 = (
    " || m003 (24-sep-2026): LAS URLs ANTERIORES YA NO SON VALIDAS (Solera re-sincronizo su "
    "catalogo en sep-2026; 8 de 11 ya no existen) y 7 productos se reemplazaron por precio y "
    "funcionalidad con aprobacion del usuario. URLs VIGENTES (SKU entre parentesis): "
    "Me01 (0549) -> https://solera.co.cr/producto/mejorante-pannivita-pro-pan-1-kg-homo-plus/ ; "
    "Ha01 (0016) -> https://solera.co.cr/producto/harina-flores-25-kg-fuerte/ ; "
    "Az01 (0021) -> https://solera.co.cr/producto/azucar-dona-maria-2kg-24-kg/ ; "
    "Sa01 (0995) -> https://solera.co.cr/producto/sal-saru-comun-economica-25-kg-saco-brin/ ; "
    "Ma02 (0030) -> https://solera.co.cr/producto/manteca-max-pan-bulk-25-kgs/ ; "
    "Le01 (0137) -> https://solera.co.cr/producto/levadura-fresca-golondrina-400-gr-cip/ ; "
    "Me02 (0957) -> https://solera.co.cr/producto/mermelada-guayaba-horneo-5-kg-ccr-caja-c2-fruta-triturada/ ; "
    "Bo01 (0516) -> https://solera.co.cr/producto/bolsas-papel-kraft-2-baguette-100-unds-biodegradable/ ; "
    "Bu01 (0179) -> https://solera.co.cr/producto/burbuja-foam-6-125-unid-postre-bl/ ; "
    "Ca01 (0474) -> https://solera.co.cr/producto/cafe-naranjo-puro-250gr-x-10-unidades/ ; "
    "Ma03 (0034) -> https://solera.co.cr/producto/margarina-tumba-numar-7-5-kilos-60-barritas-x125gr/ . "
    "Nunca usar 'SAL ROCA 25 KG' (no es para consumo humano)."
)
NOTA_N72 = (
    " || m003 (24-sep-2026): METODO ACTUALIZADO por decision del usuario. Principal = API "
    "publica de catalogo de cada tienda (gratis, sin IA, no depende del diseno): Solera "
    "(WooCommerce Store API) https://solera.co.cr/wp-json/wc/store/v1/products?slug=<slug> , "
    "precio en prices.price (currency_minor_unit 0); MaxiPali (VTEX) "
    "https://www.maxipali.co.cr/api/catalog_system/pub/products/search?fq=alternateIds_Ean:<EAN> , "
    "precio en items[0].sellers[0].commertialOffer.Price. Respaldo = ScrapeGraphAI con Google "
    "Gemini (clave GEMINI_API_KEY en el .env). Precios 1 vez al dia; si falla, NO se toca la celda."
)
NOTA_L72 = " || m003: API de catalogo (principal) + ScrapeGraphAI/Gemini (respaldo)"
NOTA_N76 = (
    " || m003 (24-sep-2026): en Ingreso de Inventario (C10, C19) y Control de Inventario "
    "(A12, A21) seguian los nombres viejos '5 KILOS' y 'X 10 UNIDADES', distintos de la tabla MP "
    "(esas compras salian con precio 0 y como 'Compras Externas Netas'). La m003 puso el mismo "
    "nombre nuevo en las 3 hojas. Ver COST.MP.FUENTE para las URLs vigentes."
)
NOTA_N232 = (
    " || m003 (24-sep-2026): D10 = 1050, precio verificado en MaxiPali el 24-sep-2026 "
    "(EAN 7441001601064, Leche Semidescremada Dos Pinos 1 L). Respaldo si no hay: Sabemas UHT 1 L "
    "(EAN 7441078217694). Ya no es placeholder; el robot lo actualizara a diario."
)
NOTAS_E = [(DIC, "N70", NOTA_N70), (DIC, "N72", NOTA_N72), (DIC, "L72", NOTA_L72),
           (DIC, "N76", NOTA_N76)]
NOTA_PRECIOS = [(DIC, "N232", NOTA_N232)]


class MigracionDetenida(Exception):
    """Una celda no tenia el valor esperado: no se escribio nada."""


def _lista_cambios(wb, con_precios=True):
    """Arma la lista final (hoja, celda, antes, despues, bloque)."""
    cambios = [(h, c, a, d, "A") for h, c, a, d in BLOQUE_A]
    cambios += [(h, c, a, d, "B") for h, c, a, d in BLOQUE_B]
    cambios += [(h, c, a, d, "C") for h, c, a, d in BLOQUE_C]
    if con_precios:
        cambios += [(h, c, a, d, "D") for h, c, a, d in BLOQUE_D]
    notas = NOTAS_E + (NOTA_PRECIOS if con_precios else [])
    for h, c, nota in notas:
        actual = wb[h][c].value
        if isinstance(actual, str) and actual.endswith(nota):
            cambios.append((h, c, actual[: -len(nota)], actual, "E"))
        else:
            cambios.append((h, c, actual, (actual or "") + nota, "E"))
    return cambios


def migrar(wb, wb_valores=None, con_precios=True):
    cambios = _lista_cambios(wb, con_precios)

    # 1) Revisar TODO antes de escribir
    problemas, pendientes = [], []
    for hoja, celda, antes, despues, bloque in cambios:
        valor = wb[hoja][celda].value
        if valor == despues:
            continue
        if valor != antes:
            problemas.append(f"'{hoja}'!{celda}: se esperaba {antes!r} y tiene {valor!r}")
        else:
            pendientes.append((hoja, celda, antes, despues, bloque))
    if problemas:
        raise MigracionDetenida("NO SE CAMBIO NADA. Celdas distintas a lo esperado:\n  "
                                + "\n  ".join(problemas))

    # 2) Regla del proyecto: ningun TEXTO nuevo empieza con "="
    for hoja, celda, antes, despues, bloque in pendientes:
        if isinstance(despues, str) and despues.startswith("=") and bloque != "C":
            raise MigracionDetenida(f"Texto que empieza con '=' en '{hoja}'!{celda}. No se escribio nada.")

    if not pendientes:
        return {"hubo_cambios": False, "cambios": ["Ya estaba aplicada: 0 celdas pendientes."],
                "detalle": []}

    # 3) Escribir
    for hoja, celda, antes, despues, bloque in pendientes:
        wb[hoja][celda] = despues

    # 4) Volver a leer y confirmar
    for hoja, celda, antes, despues, bloque in pendientes:
        if wb[hoja][celda].value != despues:
            raise MigracionDetenida(f"Verificacion fallo en '{hoja}'!{celda}")

    por_bloque = {}
    for p in pendientes:
        por_bloque[p[4]] = por_bloque.get(p[4], 0) + 1
    resumen = [f"Bloque {b}: {n} celdas" for b, n in sorted(por_bloque.items())]

    from conexion_drive.tablas import registrar_evento
    registrar_evento(
        wb, modulo="Sistema", tipo_evento="MIGRACION", entidad="Libro Maestro",
        id_entidad=ID_MIGRACION, accion="Proveedores 3C (MaxiPali/Solera)", estado="OK",
        detalle=("A) Leche Suli -> " + LECHE_NUEVA + " (11 celdas) | B) 7 materias primas de Solera "
                 "renombradas en MP, Ingreso y Control | C) 6 conversiones col E de MP"
                 + (" | D) precios verificados 24-sep (MP D10:D20, TCM D10:D12, D14)" if con_precios else "")
                 + " | E) notas '|| m003' en DiccionarioDatos | Total: "
                 + f"{len(pendientes)} celdas"),
    )
    return {"hubo_cambios": True, "cambios": resumen, "detalle": pendientes}


def _aplicar_a_archivo(ruta, con_precios):
    """Aplica la migracion a un .xlsx suelto y lo recalcula con LibreOffice."""
    import io
    from pathlib import Path
    import openpyxl
    from conexion_drive.recalculo import reparar_prefijos_xlfn, recalcular_bytes

    ruta = Path(ruta)
    contenido, _ = reparar_prefijos_xlfn(ruta.read_bytes())
    wb = openpyxl.load_workbook(io.BytesIO(contenido))
    resultado = migrar(wb, None, con_precios)
    if resultado["hubo_cambios"]:
        salida = io.BytesIO()
        wb.save(salida)
        ruta.write_bytes(recalcular_bytes(salida.getvalue()))
    return resultado


if __name__ == "__main__":
    con_precios = "--sin-precios" not in sys.argv
    try:
        if "--archivo" in sys.argv:
            ruta = sys.argv[sys.argv.index("--archivo") + 1]
            print(f"Excel de destino: ARCHIVO {ruta}")
            resultado = _aplicar_a_archivo(ruta, con_precios)
        else:
            from conexion_drive import config
            from conexion_drive.drive_excel import editar_excel
            print(f"Excel de destino: {'COPIA LOCAL ' + str(config.EXCEL_LOCAL_PATH) if config.ORIGEN_EXCEL == 'local' else 'GOOGLE DRIVE (real)'}")
            resultado = editar_excel(lambda wb, wv: migrar(wb, wv, con_precios),
                                     descripcion="Migracion m003_proveedores_3c")
    except MigracionDetenida as error:
        print(str(error))
        sys.exit(1)
    for linea in resultado["cambios"]:
        print("  " + linea)
    print("LISTO." if resultado["hubo_cambios"] else "Ya estaba aplicada: no se cambio nada.")
