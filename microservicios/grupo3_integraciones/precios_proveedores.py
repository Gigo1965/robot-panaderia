"""Precios de proveedores (Grupo 3C) - capa principal: APIs publicas de catalogo.

Que hace (Regla #18 del libro de mapas):
  Vocabulario: 2 PROVEEDORES (Solera, MaxiPali), cada uno con su API DE CATALOGO
  (fuente de datos JSON); 16 PRODUCTOS; 2 TABLAS DESTINO en el Excel maestro.

  1. Consulta el precio de cada producto en la API de catalogo de su proveedor:
       - Solera (WooCommerce Store API) por slug; si el slug ya no existe, por SKU.
       - MaxiPali (VTEX) por codigo de barras (EAN). Leche: si Dos Pinos no esta
         disponible, usa el respaldo Sabemas UHT 1 L.
  2. Escribe SOLO la columna D (precio) de 'TABLA CRUCE MAESTRA MP' (Solera) y
     'TABLA CRUCE MAESTRA' (MaxiPali). Todo lo demas lo recalcula el Excel.
  3. Si una consulta falla, NO toca esa celda (queda el precio anterior).
  4. Deja un evento en RegistroEventos con lo que cambio y lo que fallo.

Por que primero se consulta y DESPUES se abre el Excel:
  las consultas tardan ~1 minuto (pausas de cortesia). Si se hicieran con el Excel
  abierto, aumenta el riesgo de chocar con una venta que llega al mismo tiempo.
  Asi el Excel se abre solo unos segundos para escribir.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.precios_proveedores --solo-consultar
        (consulta las APIs de catalogo y muestra los precios; NO toca ningun Excel)
    py -m microservicios.grupo3_integraciones.precios_proveedores --local
        (escribe en la copia datos_prueba/maestro_prueba.xlsx)
    py -m microservicios.grupo3_integraciones.precios_proveedores
        (escribe en el Excel real de Google Drive)

Limite de cambio de precio (se activa DESPUES de la primera corrida real):
    en el .env agregar  LIMITE_CAMBIO_PRECIO=0.5   (= 50 %). Si un precio nuevo
    difiere mas que eso del anterior, no se escribe y queda marcado en el evento.
"""
import json
import os
import random
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

# --local: igual que las migraciones. Debe ir ANTES de importar conexion_drive.config.
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

# ---------------------------------------------------------------- Parametros (mapa MP, filas 5-6)
TIMEOUT_S = 30
REINTENTOS = 4
PAUSA_MIN_S, PAUSA_MAX_S = 2.5, 5.5
PRECIO_MAXIMO_RAZONABLE = 1_000_000   # un precio mayor se toma como lectura erronea
USER_AGENT = "RobotPanaderiaCanas/1.0 (proyecto academico de maestria)"

HOJA_MP = "TABLA CRUCE MAESTRA MP"     # Codigo en col F, precio en col D
HOJA_TCM = "TABLA CRUCE MAESTRA"       # Codigo en col E, precio en col D

SOLERA_API = "https://solera.co.cr/wp-json/wc/store/v1/products"
VTEX_API = "https://www.maxipali.co.cr/api/catalog_system/pub/products/search"

# Codigo -> (slug, SKU). Verificados el 24-sep-2026 (libro de mapas V10).
SOLERA = {
    "Me01": ("mejorante-pannivita-pro-pan-1-kg-homo-plus", "0549"),
    "Ha01": ("harina-flores-25-kg-fuerte", "0016"),
    "Az01": ("azucar-dona-maria-2kg-24-kg", "0021"),
    "Sa01": ("sal-saru-comun-economica-25-kg-saco-brin", "0995"),
    "Ma02": ("manteca-max-pan-bulk-25-kgs", "0030"),
    "Le01": ("levadura-fresca-golondrina-400-gr-cip", "0137"),
    "Me02": ("mermelada-guayaba-horneo-5-kg-ccr-caja-c2-fruta-triturada", "0957"),
    "Bo01": ("bolsas-papel-kraft-2-baguette-100-unds-biodegradable", "0516"),
    "Bu01": ("burbuja-foam-6-125-unid-postre-bl", "0179"),
    "Ca01": ("cafe-naranjo-puro-250gr-x-10-unidades", "0474"),
    "Ma03": ("margarina-tumba-numar-7-5-kilos-60-barritas-x125gr", "0034"),
}
# Codigo -> lista de EAN en orden de preferencia (el 2o es el respaldo).
MAXIPALI = {
    "L01": ["7441001601064", "7441078217694"],   # Dos Pinos; respaldo Sabemas UHT 1 L
    "Q01": ["7441013503189"],
    "Ma01": ["0751963110014"],
    "CC01": ["7441003500488"],
    "Na01": ["7441001608728"],
}


class ConsultaFallida(Exception):
    """La API del proveedor no devolvio un precio valido para ese producto."""


# ---------------------------------------------------------------- Red
_primera_llamada = True


def _pausa_de_cortesia():
    """Espera entre llamadas para no saturar la API del proveedor (secuencial, nunca en paralelo)."""
    global _primera_llamada
    if _primera_llamada:
        _primera_llamada = False
        return
    time.sleep(random.uniform(PAUSA_MIN_S, PAUSA_MAX_S))


def _get_json(url):
    """GET con timeout y reintentos. Devuelve el JSON o lanza ConsultaFallida."""
    ultimo_error = None
    for intento in range(1, REINTENTOS + 1):
        _pausa_de_cortesia()
        try:
            pedido = urllib.request.Request(url, headers={"User-Agent": USER_AGENT,
                                                          "Accept": "application/json"})
            with urllib.request.urlopen(pedido, timeout=TIMEOUT_S) as respuesta:
                return json.loads(respuesta.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            ultimo_error = f"HTTP {e.code}"
            if e.code in (400, 401, 403, 404):
                break          # reintentar no lo va a arreglar
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            ultimo_error = f"sin conexion ({e})"
        except json.JSONDecodeError:
            ultimo_error = "la respuesta no es JSON"
    raise ConsultaFallida(ultimo_error or "error desconocido")


def _validar(precio):
    if not isinstance(precio, (int, float)) or precio <= 0 or precio > PRECIO_MAXIMO_RAZONABLE:
        raise ConsultaFallida(f"precio fuera de rango: {precio!r}")
    return round(precio)


# ---------------------------------------------------------------- Solera (WooCommerce)
def _precio_woo(producto):
    """prices.price viene como texto en la unidad menor (currency_minor_unit decimales)."""
    precios = producto.get("prices") or {}
    decimales = int(precios.get("currency_minor_unit", 0))
    return int(precios["price"]) / (10 ** decimales)


def consultar_solera(codigo):
    slug, sku = SOLERA[codigo]
    datos = _get_json(f"{SOLERA_API}?slug={urllib.parse.quote(slug)}")
    via = "slug"
    if not datos:   # el slug cambio: buscar por SKU
        datos = [p for p in _get_json(f"{SOLERA_API}?sku={urllib.parse.quote(sku)}")
                 if str(p.get("sku")) == sku]
        via = "SKU"
    if not datos:
        raise ConsultaFallida(f"no se encontro por slug ni por SKU {sku}")
    producto = datos[0]
    try:
        precio = _validar(_precio_woo(producto))
    except (KeyError, ValueError, TypeError):
        raise ConsultaFallida("la respuesta no trae prices.price")
    nota = f"via {via}"
    if producto.get("is_in_stock") is False:
        nota += ", SIN STOCK segun el proveedor"
    return precio, nota


# ---------------------------------------------------------------- MaxiPali (VTEX)
def _precio_vtex(datos, ean):
    """Busca el item con ese EAN y devuelve (precio, disponible)."""
    for producto in datos or []:
        for item in producto.get("items", []):
            if str(item.get("ean")) == ean or len(producto.get("items", [])) == 1:
                oferta = item["sellers"][0]["commertialOffer"]
                return oferta["Price"], bool(oferta.get("IsAvailable", True))
    raise ConsultaFallida(f"EAN {ean} no encontrado")


def consultar_maxipali(codigo):
    errores = []
    for n, ean in enumerate(MAXIPALI[codigo]):
        try:
            datos = _get_json(f"{VTEX_API}?fq=alternateIds_Ean:{ean}")
            precio, disponible = _precio_vtex(datos, ean)
            precio = _validar(precio)
            if not disponible and n < len(MAXIPALI[codigo]) - 1:
                errores.append(f"EAN {ean} sin disponibilidad")
                continue            # probar el respaldo
            nota = f"EAN {ean}" + (" (RESPALDO)" if n > 0 else "")
            if not disponible:
                nota += ", SIN DISPONIBILIDAD segun el proveedor"
            return precio, nota
        except (ConsultaFallida, KeyError, IndexError, TypeError) as e:
            errores.append(f"EAN {ean}: {e}")
    raise ConsultaFallida("; ".join(errores))


# ---------------------------------------------------------------- Paso 1: consultar
def consultar_todo():
    """Devuelve lista de dicts: hoja, codigo, precio (o None), nota."""
    resultados = []
    for hoja, tabla, funcion in ((HOJA_MP, SOLERA, consultar_solera),
                                 (HOJA_TCM, MAXIPALI, consultar_maxipali)):
        for codigo in tabla:
            try:
                precio, nota = funcion(codigo)
            except ConsultaFallida as e:
                precio, nota = None, f"FALLO: {e}"
            resultados.append({"hoja": hoja, "codigo": codigo, "precio": precio, "nota": nota})
            print(f"  {codigo:<5} {('₡' + format(precio, ',')) if precio else '—':>10}  {nota}")
    return resultados


# ---------------------------------------------------------------- Paso 2: escribir
def _fila_del_codigo(ws, col_codigo, codigo):
    for fila in range(1, ws.max_row + 1):
        if str(ws[f"{col_codigo}{fila}"].value or "").strip() == codigo:
            return fila
    return None


def aplicar_precios(wb, resultados, limite=None):
    """Escribe en col D solo los precios validos que cambiaron. Devuelve el resumen."""
    cambios, iguales, fallos, frenados = [], 0, [], []
    for r in resultados:
        if r["precio"] is None:
            fallos.append(f"{r['codigo']} ({r['nota']})")
            continue
        ws = wb[r["hoja"]]
        fila = _fila_del_codigo(ws, "F" if r["hoja"] == HOJA_MP else "E", r["codigo"])
        if fila is None:
            fallos.append(f"{r['codigo']} (codigo no encontrado en '{r['hoja']}')")
            continue
        celda = f"D{fila}"
        anterior = ws[celda].value
        if isinstance(anterior, (int, float)) and round(anterior) == r["precio"]:
            iguales += 1
            continue
        if limite and isinstance(anterior, (int, float)) and anterior > 0 \
                and abs(r["precio"] - anterior) / anterior > limite:
            frenados.append(f"{r['codigo']} {anterior}->{r['precio']}")
            continue
        ws[celda] = r["precio"]
        hoja_corta = "MP" if r["hoja"] == HOJA_MP else "TCM"
        cambios.append(f"{r['codigo']} {hoja_corta}!{celda}: {anterior}->{r['precio']}")
    return {"cambios": cambios, "iguales": iguales, "fallos": fallos, "frenados": frenados}


def main():
    solo_consultar = "--solo-consultar" in sys.argv
    print("=== 1) Consultando APIs de catalogo (2 proveedores, 16 productos; uno a la vez, con pausas) ===")
    resultados = consultar_todo()
    validos = sum(1 for r in resultados if r["precio"] is not None)
    print(f"\nPrecios obtenidos: {validos} de {len(resultados)} productos")
    if solo_consultar:
        print("Modo --solo-consultar: no se toco ningun Excel.")
        return

    from conexion_drive import config
    from conexion_drive.drive_excel import editar_excel
    from conexion_drive.tablas import registrar_evento

    limite = float(os.getenv("LIMITE_CAMBIO_PRECIO") or 0) or None
    destino = ("COPIA LOCAL " + str(config.EXCEL_LOCAL_PATH)) if config.ORIGEN_EXCEL == "local" \
        else "GOOGLE DRIVE (real)"
    print(f"\n=== 2) Escribiendo en: {destino} ===")

    def editar(wb, wb_valores):
        resumen = aplicar_precios(wb, resultados, limite)
        detalle = (f"Cambiados {len(resumen['cambios'])}: " + ("; ".join(resumen["cambios"]) or "-")
                   + f" | Sin cambio: {resumen['iguales']}"
                   + f" | Fallidos (celda NO tocada) {len(resumen['fallos'])}: "
                   + ("; ".join(resumen["fallos"]) or "-")
                   + (f" | Frenados por limite {limite:.0%}: " + "; ".join(resumen["frenados"])
                      if resumen["frenados"] else ""))
        registrar_evento(
            wb, modulo="Integraciones", tipo_evento="PRECIOS_PROVEEDORES", entidad="Tablas maestras",
            id_entidad="precios_proveedores", accion="Actualizar precios (API catalogo)",
            estado="OK" if not resumen["fallos"] and not resumen["frenados"] else "PARCIAL",
            detalle=detalle, origen="Robot Python",
        )
        resumen["hubo_cambios"] = True     # siempre se guarda el evento (auditoria diaria)
        return resumen

    resumen = editar_excel(editar, descripcion="Precios de proveedores")
    for linea in resumen["cambios"]:
        print("  CAMBIO  " + linea)
    print(f"  Sin cambio: {resumen['iguales']}")
    for linea in resumen["fallos"]:
        print("  FALLO (celda no tocada)  " + linea)
    for linea in resumen["frenados"]:
        print("  FRENADO por limite  " + linea)
    print("LISTO.")


if __name__ == "__main__":
    main()
