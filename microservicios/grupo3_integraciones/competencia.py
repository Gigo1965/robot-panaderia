"""Promociones de la competencia (Grupo 3C): Musmanni y Musi.

Que hace (Reglas #18 y NODO INTEGRAL del libro de mapas):
  Vocabulario: 2 COMPETIDORES (Musmanni y Musi, ambos de FIFCO). Cada uno publica
  sus promociones en un sitio WordPress. Se escriben 2 TABLAS del Excel maestro:
  NODO INTEGRAL (resumen) y Tabla Combos Dinamicos (detalle y valor).

  Capa 1 - API estructurada, SIN IA (gratis):
    Consulta la API de WordPress de cada competidor (paginas de promociones y,
    en Musmanni, el tipo de contenido "promos"). Se queda SOLO con las imagenes
    VISIBLES: si el sitio las oculto en computadora, tablet y celular
    (marcas elementor-hidden-*), no cuentan como promocion publicada.
  Capa 2 - vision con Google Gemini, SOLO para imagenes visibles:
    El precio y la vigencia vienen DENTRO de las imagenes. Gemini devuelve un
    JSON (producto, precio, vigencia). Cada imagen se lee UNA sola vez: el
    resultado se guarda en datos_prueba/cache_competencia.json.
  Capa 3 - decision y escritura:
    VIGENTE     -> hay al menos una promo con fecha vigente.
    SIN_PROMO   -> la consulta funciono, pero no hay promo vigente:
                   se escribe "Sin promocion vigente (fecha)" y valor 0.
    FALLO       -> no hubo conexion: NO se tocan las celdas.
  Deja un evento COMPETENCIA en RegistroEventos.

Celdas destino:
    Musmanni: NODO INTEGRAL C11 (resumen) | Tabla Combos Dinamicos D10 (detalle), E10 (valor)
    Musi:     NODO INTEGRAL C21 (resumen) | Tabla Combos Dinamicos D11 (detalle), E11 (valor)
              (las celdas de Musi las crea la migracion m004; sin ella el robot se detiene)

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.competencia --solo-consultar --sin-vision
        (solo capa 1: lista las imagenes visibles/ocultas; no usa Gemini ni Excel)
    py -m microservicios.grupo3_integraciones.competencia --probar-vision <URL de imagen>
        (lee UNA imagen con Gemini y muestra el resultado; no usa cache ni Excel)
    py -m microservicios.grupo3_integraciones.competencia --solo-consultar
        (capas 1 y 2 completas; NO toca ningun Excel)
    py -m microservicios.grupo3_integraciones.competencia --local
        (escribe en la copia datos_prueba/maestro_prueba.xlsx)
    py -m microservicios.grupo3_integraciones.competencia
        (escribe en el Excel real de Google Drive)

La clave de Gemini se lee SIEMPRE del .env (GEMINI_API_KEY); nunca va en el codigo.
Modelo: opcional GEMINI_MODELO en el .env; si no esta, se elige el "flash" mas nuevo.
"""
import base64
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from html.parser import HTMLParser
from pathlib import Path

# --local: igual que las migraciones. Debe ir ANTES de importar conexion_drive.config.
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import config  # noqa: E402  (carga el .env: GEMINI_API_KEY)

# ---------------------------------------------------------------- Parametros
TIMEOUT_S = 30
REINTENTOS = 3
PAUSA_WEB_S = 2.0                 # entre llamadas al sitio del competidor
PAUSA_GEMINI_S = 7.0              # nivel gratuito de Gemini: pocas llamadas por minuto
DIAS_PROMOS_API = 60              # una "promo" de la API mas vieja que esto no se considera
DIAS_SIN_VIGENCIA = 45            # imagen sin fecha impresa: vigente solo si se subio hace <= 45 dias
PRECIO_MAXIMO_RAZONABLE = 100_000 # una promo de panaderia mayor a esto se toma como lectura erronea
USER_AGENT = "RobotPanaderiaCanas/1.0 (proyecto academico de maestria)"
API_GEMINI = "https://generativelanguage.googleapis.com/v1beta"
# Google retira modelos cada cierto tiempo (25-sep-2026: gemini-2.5-flash ya respondia 404).
# Por eso el modelo NO va fijo: si el .env trae GEMINI_MODELO se usa ese; si no, el robot
# pregunta a Google que modelos "flash" tiene la clave y elige el mas nuevo.
CACHE = config.RAIZ_PROYECTO / "datos_prueba" / "cache_competencia.json"
HORA_CR = timezone(timedelta(hours=-6))   # Costa Rica: UTC-6 todo el ano (sin horario de verano)

HOJA_NODO = "NODO INTEGRAL"
HOJA_COMBOS = "Tabla Combos Dinamicos"

COMPETIDORES = {
    "Musmanni": {
        "base": "https://musmanni.com",
        "paginas": ["promociones", "promociones-reposteria",
                    "promociones-almuerzos", "promociones-pasteles"],
        "tipo_promos": "promos",               # /wp-json/wp/v2/promos
        "celda_nodo": "C11", "celda_detalle": "D10", "celda_valor": "E10",
        "rotulo_requerido": None,
    },
    "Musi": {
        "base": "https://tiendasmusi.com",
        "paginas": ["promociones-para-panaderia-y-reposteria"],
        "tipo_promos": None,
        "celda_nodo": "C21", "celda_detalle": "D11", "celda_valor": "E11",
        "rotulo_requerido": ("B21", "DATA OFERTA MUSI"),   # lo crea la migracion m004
    },
}

PROMPT = """Eres analista de precios de la competencia de una panaderia en Costa Rica.
La imagen fue publicada en el sitio web de {competidor} y subida en {fecha_subida}.
Responde SOLO un objeto JSON con estas llaves:
  "es_promocion": true si la imagen anuncia una oferta, combo o descuento con precio;
                  false si es publicidad general, un logo o un banner sin oferta.
  "producto": texto corto con lo que incluye la oferta (si hay varias, nombralas todas).
  "precio_colones": numero entero sin simbolos (si hay varias ofertas, el precio mas bajo),
                    o null si no aparece un precio.
  "vigencia_desde": fecha AAAA-MM-DD o null si la imagen no la indica.
  "vigencia_hasta": fecha AAAA-MM-DD o null si la imagen no la indica.
                    Si la imagen no dice el ano, usa el ano de subida.
  "texto_vigencia": el texto de vigencia tal como aparece en la imagen, o null.
No inventes datos que no esten en la imagen."""


class ConsultaFallida(Exception):
    """El sitio o la API no respondio de forma util."""


def hoy_cr():
    return datetime.now(HORA_CR).date()


# ---------------------------------------------------------------- Red
def _pedir(url, datos=None, encabezados=None):
    """GET (o POST si hay datos) con timeout y reintentos. Devuelve (bytes, tipo_mime)."""
    ultimo_error = None
    for intento in range(1, REINTENTOS + 1):
        try:
            cabeceras = {"User-Agent": USER_AGENT, **(encabezados or {})}
            pedido = urllib.request.Request(url, data=datos, headers=cabeceras)
            with urllib.request.urlopen(pedido, timeout=TIMEOUT_S) as respuesta:
                tipo = (respuesta.headers.get("Content-Type") or "").split(";")[0].strip()
                return respuesta.read(), tipo
        except urllib.error.HTTPError as e:
            try:
                mensaje = json.loads(e.read().decode("utf-8"))["error"]["message"]
                ultimo_error = f"HTTP {e.code} - {mensaje[:200]}"
            except Exception:
                ultimo_error = f"HTTP {e.code}"
            if e.code in (429, 500, 502, 503, 504):
                time.sleep(20 * intento)   # saturado: esperar mas y reintentar
                continue
            break                           # 400/401/403/404: reintentar no lo arregla
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            ultimo_error = f"sin conexion ({e})"
            time.sleep(3 * intento)
    raise ConsultaFallida(ultimo_error or "error desconocido")


def _get_json(url):
    time.sleep(PAUSA_WEB_S)
    contenido, _ = _pedir(url, encabezados={"Accept": "application/json"})
    try:
        return json.loads(contenido.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        raise ConsultaFallida("la respuesta no es JSON")


# ---------------------------------------------------------------- Capa 1: imagenes visibles
DISPOSITIVOS = ("desktop", "tablet", "mobile")
ETIQUETAS_VACIAS = {"img", "br", "hr", "meta", "link", "input", "source", "wbr",
                    "area", "base", "col", "embed", "param", "track"}


class _LectorImagenes(HTMLParser):
    """Recorre el HTML y anota cada imagen con los dispositivos donde esta oculta.

    Elementor oculta un bloque con las clases elementor-hidden-desktop / -tablet /
    -mobile. Si un contenedor esta oculto, todo lo que tiene adentro tambien.
    Las imagenes "perezosas" (lazy load) traen la URL real en <noscript> o en
    data-src; los rellenos "data:image/..." se ignoran.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.pila = []          # [(etiqueta, conjunto_de_dispositivos_ocultos)]
        self.imagenes = []      # [(url, visible)]

    def _ocultos_actuales(self):
        return self.pila[-1][1] if self.pila else frozenset()

    def _ocultos(self, attrs):
        clases = (dict(attrs).get("class") or "").split()
        propios = {d for d in DISPOSITIVOS if f"elementor-hidden-{d}" in clases}
        return self._ocultos_actuales() | frozenset(propios)

    def _anotar_img(self, attrs, ocultos):
        d = dict(attrs)
        for llave in ("data-src", "data-lazy-src", "src", "data-srcset", "srcset"):
            valor = (d.get(llave) or "").strip()
            if valor and not valor.startswith("data:"):
                url = valor.split(",")[0].split()[0]
                self.imagenes.append((url, len(ocultos) < len(DISPOSITIVOS)))
                return

    def handle_starttag(self, tag, attrs):
        ocultos = self._ocultos(attrs)
        if tag == "img":
            self._anotar_img(attrs, ocultos)
        if tag not in ETIQUETAS_VACIAS:
            self.pila.append((tag, ocultos))

    def handle_startendtag(self, tag, attrs):
        if tag == "img":
            self._anotar_img(attrs, self._ocultos(attrs))

    def handle_endtag(self, tag):
        if tag in ETIQUETAS_VACIAS:
            return
        for i in range(len(self.pila) - 1, -1, -1):
            if self.pila[i][0] == tag:
                del self.pila[i:]
                return


def _canonica(url):
    """Quita el tamano que agrega WordPress (Promo-1-1-1024x1024.png -> Promo-1-1.png)."""
    return re.sub(r"-\d+x\d+(?=\.\w+$)", "", url.split("?")[0])


def fecha_subida(url):
    """WordPress guarda las imagenes en /uploads/AAAA/MM/: sirve como fecha aproximada."""
    m = re.search(r"/uploads/(\d{4})/(\d{2})/", url)
    return date(int(m.group(1)), int(m.group(2)), 1) if m else None


def imagenes_del_html(html):
    """Devuelve lista ordenada de dicts {url, visible} sin repetidos (solo fotos de uploads)."""
    lector = _LectorImagenes()
    lector.feed(html)
    resultado = {}
    for url, visible in lector.imagenes:
        nombre_archivo = url.lower().rsplit("/", 1)[-1]
        if "/wp-content/uploads/" not in url or nombre_archivo.endswith(".svg") \
                or "logo" in nombre_archivo:
            continue       # logos y titulos no son promociones (ahorra llamadas a Gemini)
        clave = _canonica(url)
        if clave not in resultado:
            resultado[clave] = {"url": url, "visible": visible}
        elif visible and not resultado[clave]["visible"]:
            resultado[clave] = {"url": url, "visible": True}
    return list(resultado.values())


def consultar_paginas(nombre, datos):
    """Capa 1 de un competidor. Devuelve (imagenes, notas). Lanza ConsultaFallida si TODO fallo."""
    imagenes, notas, respondio = [], [], False
    for slug in datos["paginas"]:
        url_api = (f"{datos['base']}/wp-json/wp/v2/pages?slug={urllib.parse.quote(slug)}"
                   "&_fields=id,link,modified,content")
        try:
            paginas = _get_json(url_api)
            respondio = True
        except ConsultaFallida as e:
            notas.append(f"{slug}: FALLO ({e})")
            continue
        if not paginas:
            notas.append(f"{slug}: la pagina ya no existe")
            continue
        pagina = paginas[0]
        html = (pagina.get("content") or {}).get("rendered") or ""
        encontradas = imagenes_del_html(html)
        if not encontradas:        # la API no trajo el diseno completo: leer la pagina publica
            try:
                time.sleep(PAUSA_WEB_S)
                contenido, _ = _pedir(pagina["link"])
                encontradas = imagenes_del_html(contenido.decode("utf-8", errors="replace"))
            except ConsultaFallida as e:
                notas.append(f"{slug}: FALLO al leer la pagina publica ({e})")
        visibles = sum(1 for i in encontradas if i["visible"])
        notas.append(f"{slug}: modificada {str(pagina.get('modified', '?'))[:10]}, "
                     f"{visibles} imagen(es) visible(s), {len(encontradas) - visibles} oculta(s)")
        for img in encontradas:
            img["origen"] = slug
            img["fecha"] = fecha_subida(img["url"])
        imagenes.extend(encontradas)

    if datos["tipo_promos"]:
        respondio = _consultar_tipo_promos(datos, imagenes, notas) or respondio

    if not respondio:
        raise ConsultaFallida("ninguna consulta respondio")
    return imagenes, notas


def _consultar_tipo_promos(datos, imagenes, notas):
    """Musmanni tiene un tipo de contenido 'promos' en su API. Solo cuentan las recientes."""
    url = (f"{datos['base']}/wp-json/wp/v2/{datos['tipo_promos']}"
           "?per_page=20&_fields=id,date,link,title,featured_media")
    try:
        promos = _get_json(url)
    except ConsultaFallida as e:
        notas.append(f"API {datos['tipo_promos']}: FALLO ({e})")
        return False
    limite = hoy_cr() - timedelta(days=DIAS_PROMOS_API)
    recientes = 0
    for promo in promos or []:
        publicada = datetime.fromisoformat(promo["date"]).date()
        if publicada < limite or not promo.get("featured_media"):
            continue
        try:
            media = _get_json(f"{datos['base']}/wp-json/wp/v2/media/{promo['featured_media']}"
                              "?_fields=source_url")
        except ConsultaFallida:
            continue
        recientes += 1
        imagenes.append({"url": media["source_url"], "visible": True,
                         "origen": f"API {datos['tipo_promos']}", "fecha": publicada})
    ultima = max((p["date"][:10] for p in promos or []), default="ninguna")
    notas.append(f"API {datos['tipo_promos']}: {len(promos or [])} promo(s), "
                 f"ultima publicada {ultima}, {recientes} de los ultimos {DIAS_PROMOS_API} dias")
    return True


# ---------------------------------------------------------------- Capa 2: vision con Gemini
_ultima_llamada_gemini = 0.0
_modelo_elegido = None


def _version(nombre):
    """'models/gemini-3.1-flash' -> (3, 1). Sirve para ordenar del mas nuevo al mas viejo."""
    m = re.search(r"gemini-(\d+)(?:\.(\d+))?", nombre)
    return (int(m.group(1)), int(m.group(2) or 0)) if m else (0, 0)


def modelo_gemini(clave):
    """Nombre del modelo a usar (ej. 'gemini-3-flash'). Se decide una vez por corrida."""
    global _modelo_elegido
    if _modelo_elegido:
        return _modelo_elegido
    fijo = (os.getenv("GEMINI_MODELO") or "").strip()
    if fijo:
        _modelo_elegido = fijo.replace("models/", "")
        return _modelo_elegido
    contenido, _ = _pedir(f"{API_GEMINI}/models?pageSize=200",
                          encabezados={"x-goog-api-key": clave})
    modelos = json.loads(contenido.decode("utf-8")).get("models", [])
    candidatos = [m["name"] for m in modelos
                  if "generateContent" in m.get("supportedGenerationMethods", [])
                  and "flash" in m["name"]
                  and not any(x in m["name"] for x in ("lite", "image", "tts", "audio",
                                                       "live", "thinking", "exp"))]
    if not candidatos:
        raise ConsultaFallida("la clave no tiene ningun modelo 'flash' disponible")
    # El mas nuevo primero; a igual version, el estable antes que el 'preview'
    candidatos.sort(key=lambda n: (_version(n), "preview" not in n), reverse=True)
    _modelo_elegido = candidatos[0].replace("models/", "")
    print(f"    (modelo Gemini elegido: {_modelo_elegido})")
    return _modelo_elegido


def leer_imagen_con_gemini(url, competidor, fecha):
    """Descarga la imagen y le pide a Gemini el JSON de la promocion."""
    global _ultima_llamada_gemini
    clave = os.getenv("GEMINI_API_KEY")
    if not clave:
        raise ConsultaFallida("falta GEMINI_API_KEY en el archivo .env")

    imagen, tipo = _pedir(url)
    if not tipo.startswith("image/"):
        tipo = "image/png" if url.lower().endswith(".png") else "image/jpeg"

    espera = PAUSA_GEMINI_S - (time.time() - _ultima_llamada_gemini)
    if espera > 0:
        time.sleep(espera)
    cuerpo = {
        "contents": [{"parts": [
            {"inline_data": {"mime_type": tipo, "data": base64.b64encode(imagen).decode("ascii")}},
            {"text": PROMPT.format(competidor=competidor,
                                   fecha_subida=fecha.isoformat() if fecha else "fecha desconocida")},
        ]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0},
    }
    url_modelo = f"{API_GEMINI}/models/{modelo_gemini(clave)}:generateContent"
    respuesta, _ = _pedir(url_modelo, datos=json.dumps(cuerpo).encode("utf-8"),
                          encabezados={"Content-Type": "application/json",
                                       "x-goog-api-key": clave})
    _ultima_llamada_gemini = time.time()
    try:
        texto = json.loads(respuesta)["candidates"][0]["content"]["parts"][0]["text"]
        leido = json.loads(texto)
    except (KeyError, IndexError, json.JSONDecodeError):
        raise ConsultaFallida("Gemini no devolvio un JSON valido")
    if isinstance(leido, list):          # a veces responde una lista con un solo objeto
        leido = leido[0] if leido else {}
    return leido


def _fecha(texto):
    try:
        return date.fromisoformat(str(texto)[:10])
    except (TypeError, ValueError):
        return None


def evaluar(leido, fecha_subida_img, hoy):
    """Decide si una promo leida por Gemini esta vigente HOY. Devuelve (estado, precio)."""
    if not leido.get("es_promocion"):
        return "NO ES PROMO", None
    precio = leido.get("precio_colones")
    if not isinstance(precio, (int, float)) or not 0 < precio <= PRECIO_MAXIMO_RAZONABLE:
        precio = None
    else:
        precio = round(precio)
    desde, hasta = _fecha(leido.get("vigencia_desde")), _fecha(leido.get("vigencia_hasta"))
    if hasta:
        if hasta < hoy:
            return f"VENCIDA (hasta {hasta.isoformat()})", precio
        if desde and desde > hoy:
            return f"FUTURA (desde {desde.isoformat()})", precio
        return f"VIGENTE (hasta {hasta.isoformat()})", precio
    if fecha_subida_img and (hoy - fecha_subida_img).days <= DIAS_SIN_VIGENCIA:
        return "VIGENTE (sin fecha impresa, subida reciente)", precio
    antiguedad = f"subida {fecha_subida_img.isoformat()[:7]}" if fecha_subida_img else "sin fecha"
    return f"VENCIDA (sin fecha impresa, {antiguedad})", precio


def _cargar_cache():
    try:
        return json.loads(CACHE.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _guardar_cache(cache):
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")


# ---------------------------------------------------------------- Pasos 1 y 2: consultar
def consultar_competidor(nombre, datos, con_vision, cache, hoy):
    """Devuelve dict con estado VIGENTE / SIN_PROMO / FALLO, promos vigentes y notas."""
    print(f"\n--- {nombre} ({datos['base']}) ---")
    try:
        imagenes, notas = consultar_paginas(nombre, datos)
    except ConsultaFallida as e:
        print(f"  FALLO de conexion: {e}  -> sus celdas NO se tocan")
        return {"competidor": nombre, "estado": "FALLO", "vigentes": [], "notas": [str(e)]}
    for nota in notas:
        print("  " + nota)

    visibles = [i for i in imagenes if i["visible"]]
    for img in imagenes:
        marca = "VISIBLE" if img["visible"] else "oculta "
        print(f"    [{marca}] {img['url']}")

    vigentes, errores_vision = [], []
    if con_vision:
        for img in visibles:
            clave = _canonica(img["url"])
            if clave in cache:
                leido, origen = cache[clave]["leido"], "cache"
            else:
                try:
                    leido = leer_imagen_con_gemini(img["url"], nombre, img["fecha"])
                except ConsultaFallida as e:
                    errores_vision.append(f"{img['url']}: {e}")
                    print(f"    Gemini FALLO en {img['url']}: {e}")
                    continue
                cache[clave] = {"leido": leido, "fecha_lectura": hoy.isoformat()}
                origen = "Gemini"
            estado, precio = evaluar(leido, img["fecha"], hoy)
            precio_txt = f"₡{precio:,}" if precio else "sin precio"
            print(f"    -> {estado} | {leido.get('producto') or '-'} | {precio_txt} ({origen})")
            if estado.startswith("VIGENTE"):
                vigentes.append({"producto": leido.get("producto") or "promocion",
                                 "precio": precio, "estado": estado})

    if con_vision and visibles and len(errores_vision) == len(visibles):
        # Gemini fallo en TODAS: no se puede afirmar que no hay promos -> no tocar celdas
        return {"competidor": nombre, "estado": "FALLO", "vigentes": [],
                "notas": notas + [f"vision: {e}" for e in errores_vision]}
    estado = "VIGENTE" if vigentes else "SIN_PROMO"
    return {"competidor": nombre, "estado": estado, "vigentes": vigentes,
            "notas": notas + [f"vision FALLO: {e}" for e in errores_vision],
            "visibles": len(visibles), "total": len(imagenes)}


# ---------------------------------------------------------------- Paso 3: escribir
def textos_para_excel(resultado, hoy):
    """(resumen NODO, detalle Tabla Combos, valor) de un competidor con estado != FALLO."""
    fecha_txt = hoy.strftime("%d-%m-%Y")
    if resultado["estado"] == "SIN_PROMO":
        texto = f"Sin promocion vigente ({fecha_txt})"
        return texto, texto, 0
    partes = []
    for p in resultado["vigentes"]:
        precio = f"₡{p['precio']:,}" if p["precio"] else "precio no visible"
        partes.append(f"{p['producto']} - {precio} [{p['estado']}]")
    precios = [p["precio"] for p in resultado["vigentes"] if p["precio"]]
    resumen = f"{len(resultado['vigentes'])} promocion(es) vigente(s) ({fecha_txt})"
    return resumen, " | ".join(partes)[:1000], min(precios) if precios else 0


def aplicar_resultados(wb, resultados, hoy):
    """Escribe las celdas de cada competidor (salvo FALLO). Devuelve lista de cambios."""
    nodo, combos = wb[HOJA_NODO], wb[HOJA_COMBOS]
    # Primero se revisa que existan TODAS las celdas; recien despues se escribe (todo o nada).
    for r in resultados:
        requerido = COMPETIDORES[r["competidor"]]["rotulo_requerido"]
        if r["estado"] != "FALLO" and requerido:
            celda, esperado = requerido
            if str(nodo[celda].value or "").strip() != esperado:
                raise RuntimeError(
                    f"Falta la celda de {r['competidor']}: NODO INTEGRAL {celda} deberia decir "
                    f"'{esperado}'. Aplica primero la migracion m004.")
    cambios = []
    for r in resultados:
        if r["estado"] == "FALLO":
            continue
        datos = COMPETIDORES[r["competidor"]]
        resumen, detalle, valor = textos_para_excel(r, hoy)
        for ws, celda, nuevo in ((nodo, datos["celda_nodo"], resumen),
                                 (combos, datos["celda_detalle"], detalle),
                                 (combos, datos["celda_valor"], valor)):
            anterior = ws[celda].value
            if anterior != nuevo:
                ws[celda] = nuevo
                cambios.append(f"{r['competidor']} {ws.title}!{celda}: {anterior!r} -> {nuevo!r}")
    return cambios


def main():
    args = sys.argv[1:]
    hoy = hoy_cr()

    if "--probar-vision" in args:
        i = args.index("--probar-vision")
        if i + 1 >= len(args):
            print("Falta la URL de la imagen. Ejemplo:\n  py -m microservicios.grupo3_integraciones."
                  "competencia --probar-vision https://sitio.com/wp-content/uploads/2025/06/x.jpg")
            return
        url = args[i + 1]
        print(f"=== Prueba de vision con Gemini ===\nImagen: {url}")
        leido = leer_imagen_con_gemini(url, "competidor", fecha_subida(url))
        print("Respuesta de Gemini:\n" + json.dumps(leido, ensure_ascii=False, indent=2))
        estado, precio = evaluar(leido, fecha_subida(url), hoy)
        print(f"Decision del robot para hoy {hoy.isoformat()}: {estado}"
              + (f" | precio ₡{precio:,}" if precio else ""))
        print("Modo --probar-vision: no se toco ningun Excel ni el cache.")
        return

    con_vision = "--sin-vision" not in args
    solo_consultar = "--solo-consultar" in args or not con_vision
    print(f"=== 1) Consultando 2 competidores (hoy {hoy.isoformat()}; "
          f"vision {'SI' if con_vision else 'NO'}) ===")
    cache = _cargar_cache()
    resultados = [consultar_competidor(n, d, con_vision, cache, hoy)
                  for n, d in COMPETIDORES.items()]
    if con_vision:
        _guardar_cache(cache)

    print("\n=== Resumen ===")
    for r in resultados:
        if r["estado"] == "FALLO":
            print(f"  {r['competidor']:<9} FALLO -> celdas no se tocan")
        elif not con_vision:
            print(f"  {r['competidor']:<9} conexion OK: {r['visibles']} imagen(es) visible(s) "
                  f"de {r['total']} (sin vision no se decide vigencia)")
        else:
            resumen, detalle, valor = textos_para_excel(r, hoy)
            print(f"  {r['competidor']:<9} {r['estado']}: {detalle} | valor {valor}")
    if solo_consultar:
        print("Modo solo consulta: no se toco ningun Excel.")
        return

    from conexion_drive.drive_excel import editar_excel
    from conexion_drive.tablas import registrar_evento

    destino = ("COPIA LOCAL " + str(config.EXCEL_LOCAL_PATH)) if config.ORIGEN_EXCEL == "local" \
        else "GOOGLE DRIVE (real)"
    print(f"\n=== 2) Escribiendo en: {destino} ===")

    def editar(wb, wb_valores):
        cambios = aplicar_resultados(wb, resultados, hoy)
        estados = "; ".join(f"{r['competidor']}: {r['estado']}" for r in resultados)
        detalle = (f"{estados} | Cambios {len(cambios)}: " + ("; ".join(cambios) or "-")
                   + " | Notas: " + " / ".join(n for r in resultados for n in r["notas"]))
        registrar_evento(
            wb, modulo="Integraciones", tipo_evento="COMPETENCIA", entidad="Competencia",
            id_entidad="competencia", accion="Revisar promociones de competidores",
            estado="OK" if all(r["estado"] != "FALLO" for r in resultados) else "PARCIAL",
            detalle=detalle[:2000], origen="Robot Python",
        )
        return {"cambios": cambios, "hubo_cambios": True}   # siempre se guarda el evento

    resumen = editar_excel(editar, descripcion="Promociones de la competencia")
    for linea in resumen["cambios"] or ["(sin cambios en las celdas; solo se registro el evento)"]:
        print("  " + linea)
    print("LISTO.")


if __name__ == "__main__":
    main()
