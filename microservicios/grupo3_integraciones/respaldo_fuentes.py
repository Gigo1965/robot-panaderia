"""Respaldo manual de fuentes: Solera y Musi (3-oct-2026, Decisiones Pendientes #56).

Por que existe (evidencia del 3-oct-2026):
  Los sitios de Solera y de Musi le piden una verificacion contra robots ("sgcaptcha") a
  los servidores de nube (GitHub Actions y Render). Desde una conexion normal (casa u
  oficina) entregan su catalogo sin problema. Entonces, en la nube, la tarea diaria no
  puede leerlos directo.

Que es el respaldo:
  Una persona de la administracion abre en su navegador la direccion del catalogo de cada
  fuente, guarda la pagina con Ctrl + S y arrastra el archivo a una carpeta de Google
  Drive. Nadie digita nada: el archivo es el mismo dato que entrega el sitio.
    respaldo_solera...json  -> catalogo de Solera (los 11 productos, por SKU)
    respaldo_musi...json    -> pagina de promociones de Musi (trae las imagenes)
  Cuando la lectura directa falla, el robot usa el archivo MAS NUEVO de la carpeta.
  Si el respaldo tiene mas de 10 dias se usa igual, pero queda marcado como VIEJO.

Direcciones que se guardan con Ctrl + S (una vez cada 10 dias o antes):
  Solera:
    https://solera.co.cr/wp-json/wc/store/v1/products?sku=0549,0016,0021,0995,0030,0137,0957,0516,0179,0474,0034&per_page=20
  Musi:
    https://tiendasmusi.com/wp-json/wp/v2/pages?slug=promociones-para-panaderia-y-reposteria&_fields=id,link,modified,content

Donde vive (lo decide este modulo):
  - Uso normal (Drive): la carpeta cuyo ID esta en RESPALDO_FUENTES_FOLDER_ID
    (.env en la PC; "Secrets" en GitHub Actions). Se acepta el ID o el enlace de la carpeta.
    La carpeta debe estar compartida con la cuenta de servicio del robot.
  - Con --local: la carpeta datos_prueba/respaldo_fuentes/ de la PC.
  - Si no hay ninguna de las dos, no hay respaldo: el robot se comporta como antes
    (lectura fallida = la celda no se toca, Regla #19).

Este modulo SOLO LEE: nunca escribe ni borra nada en la carpeta.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.respaldo_fuentes --ver --local
    py -m microservicios.grupo3_integraciones.respaldo_fuentes --ver
        (muestra que respaldo se usaria de cada fuente, su fecha y su antiguedad)
"""
import json
import os
import re
import sys
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

# --local: igual que las migraciones. Debe ir ANTES de importar conexion_drive.config.
if "--local" in sys.argv:
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ.setdefault("EXCEL_LOCAL_PATH", "datos_prueba/maestro_prueba.xlsx")

from conexion_drive import config  # noqa: E402  (carga el .env: RESPALDO_FUENTES_FOLDER_ID)
from conexion_drive.tiempo import ahora_local, zona  # noqa: E402

PREFIJOS = {"solera": "respaldo_solera", "musi": "respaldo_musi"}
CARPETA_LOCAL = "datos_prueba/respaldo_fuentes"
DIAS_AVISO = 10                  # decision del usuario (3-oct-2026): mas viejo = se usa, pero se marca
TAMANO_MAXIMO = 5 * 1024 * 1024  # un respaldo real pesa menos de 100 KB


class RespaldoNoDisponible(Exception):
    """No hay carpeta de respaldos, o no trae un archivo valido de esa fuente."""


# ---------------------------------------------------------------- Donde vive
def id_carpeta():
    """ID de la carpeta de Drive. Acepta el ID solo o el enlace completo de la carpeta."""
    valor = (os.getenv("RESPALDO_FUENTES_FOLDER_ID") or "").strip()
    m = re.search(r"/folders/([A-Za-z0-9_-]+)", valor)
    return m.group(1) if m else valor


def carpeta_local():
    valor = (os.getenv("RESPALDO_FUENTES_LOCAL_PATH") or CARPETA_LOCAL).strip()
    ruta = Path(valor)
    return ruta if ruta.is_absolute() else config.RAIZ_PROYECTO / ruta


def modo():
    """'drive', 'local' o None (None = no hay respaldo configurado)."""
    if config.ORIGEN_EXCEL == "local":
        return "local" if carpeta_local().is_dir() else None
    return "drive" if id_carpeta() else None


def dias_aviso():
    try:
        return int(os.getenv("RESPALDO_DIAS_AVISO") or DIAS_AVISO)
    except ValueError:
        return DIAS_AVISO


# ---------------------------------------------------------------- Validar el contenido
def _validar_solera(datos):
    """El catalogo de Solera: lista de productos, cada uno con sku y prices.price."""
    if not isinstance(datos, list) or not datos:
        raise ValueError("no es una lista de productos")
    buenos = [p for p in datos if isinstance(p, dict) and p.get("sku")
              and isinstance(p.get("prices"), dict) and p["prices"].get("price") is not None]
    if not buenos:
        raise ValueError("ningun producto trae sku y prices.price")
    return f"{len(buenos)} producto(s)"


def _validar_musi(datos):
    """La pagina de Musi: lista con al menos una pagina que trae content.rendered."""
    if not isinstance(datos, list) or not datos:
        raise ValueError("no es una lista de paginas")
    buenas = [p for p in datos if isinstance(p, dict)
              and isinstance((p.get("content") or {}).get("rendered"), str)]
    if not buenas:
        raise ValueError("ninguna pagina trae content.rendered")
    return f"{len(buenas)} pagina(s)"


VALIDADORES = {"solera": _validar_solera, "musi": _validar_musi}


def _interpretar(fuente, contenido):
    """bytes -> (datos, resumen). Lanza ValueError si no es el archivo esperado."""
    if len(contenido) > TAMANO_MAXIMO:
        raise ValueError("pesa demasiado para ser un respaldo")
    try:
        datos = json.loads(contenido.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        inicio = contenido[:300].decode("utf-8", errors="replace").lower()
        if "sgcaptcha" in inicio or "<html" in inicio:
            raise ValueError("es una pagina web, no el catalogo (se guardo otra pantalla)")
        raise ValueError("no es un archivo JSON")
    return datos, VALIDADORES[fuente](datos)


# ---------------------------------------------------------------- Listar candidatos
def _candidatos_local(prefijo):
    """[(nombre, momento_local, funcion_que_devuelve_bytes)] del mas nuevo al mas viejo."""
    archivos = [a for a in carpeta_local().iterdir()
                if a.is_file() and a.name.lower().startswith(prefijo)]
    archivos.sort(key=lambda a: a.stat().st_mtime, reverse=True)
    return [(a.name, datetime.fromtimestamp(a.stat().st_mtime, zona()).replace(tzinfo=None),
             a.read_bytes) for a in archivos]


def _a_hora_local(texto_utc):
    """'2026-10-03T18:11:00.000Z' (Drive, UTC) -> datetime en hora de Costa Rica."""
    momento = datetime.fromisoformat(texto_utc.replace("Z", "+00:00"))
    if momento.tzinfo is None:
        momento = momento.replace(tzinfo=timezone.utc)
    return momento.astimezone(zona()).replace(tzinfo=None)


def _candidatos_drive(prefijo):
    from conexion_drive.drive_excel import _explicar_error_google, _servicio_drive
    from googleapiclient.http import MediaIoBaseDownload
    try:
        respuesta = _servicio_drive().files().list(
            q=f"'{id_carpeta()}' in parents and trashed = false",
            fields="files(id,name,mimeType,modifiedTime,size)",
            orderBy="modifiedTime desc", pageSize=200,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    except Exception as error:
        raise RespaldoNoDisponible(
            "no pude abrir la carpeta de respaldos en Drive: "
            + _explicar_error_google(error).replace("EXCEL_MAESTRO_FILE_ID",
                                                    "RESPALDO_FUENTES_FOLDER_ID")) from error

    def descargador(id_archivo):
        def bajar():
            solicitud = _servicio_drive().files().get_media(fileId=id_archivo,
                                                            supportsAllDrives=True)
            memoria = BytesIO()
            descarga = MediaIoBaseDownload(memoria, solicitud)
            terminado = False
            while not terminado:
                _, terminado = descarga.next_chunk()
            return memoria.getvalue()
        return bajar

    candidatos = []
    for archivo in respuesta.get("files", []):
        if not archivo.get("name", "").lower().startswith(prefijo):
            continue
        if archivo.get("mimeType", "").startswith("application/vnd.google-apps"):
            continue      # un documento de Google (no un archivo subido): no es el respaldo
        candidatos.append((archivo["name"], _a_hora_local(archivo["modifiedTime"]),
                           descargador(archivo["id"])))
    candidatos.sort(key=lambda c: c[1], reverse=True)
    return candidatos


# ---------------------------------------------------------------- Leer el respaldo
_leidos = {}        # un respaldo se baja una sola vez por corrida


def leer(fuente):
    """Devuelve el respaldo valido MAS NUEVO de la fuente ('solera' o 'musi'):
    {datos, nombre, fecha (date), dias, viejo, origen, resumen, descartados}.
    Lanza RespaldoNoDisponible si no hay carpeta o ningun archivo sirve."""
    if fuente in _leidos:
        return _leidos[fuente]
    prefijo = PREFIJOS[fuente]
    donde = modo()
    if donde is None:
        raise RespaldoNoDisponible("no hay carpeta de respaldos configurada "
                                   "(RESPALDO_FUENTES_FOLDER_ID)")
    if donde == "local":
        candidatos, origen = _candidatos_local(prefijo), f"COPIA LOCAL {carpeta_local()}"
    else:
        candidatos, origen = _candidatos_drive(prefijo), "GOOGLE DRIVE (carpeta de respaldos)"
    descartados = []
    for nombre, momento, bajar in candidatos:
        try:
            datos, resumen = _interpretar(fuente, bajar())
        except ValueError as e:
            descartados.append(f"{nombre}: {e}")
            continue
        except Exception as e:       # noqa: BLE001  (un archivo que no baja no detiene la busqueda)
            descartados.append(f"{nombre}: no se pudo descargar ({type(e).__name__})")
            continue
        dias = max((ahora_local().date() - momento.date()).days, 0)
        _leidos[fuente] = {
            "fuente": fuente, "datos": datos, "nombre": nombre, "fecha": momento.date(),
            "dias": dias, "viejo": dias > dias_aviso(), "origen": origen,
            "resumen": resumen, "descartados": descartados,
        }
        return _leidos[fuente]
    if descartados:
        raise RespaldoNoDisponible(f"ningun archivo '{prefijo}...' sirve: " + "; ".join(descartados))
    raise RespaldoNoDisponible(f"no hay ningun archivo '{prefijo}...' en la carpeta de respaldos "
                               "(o la carpeta no esta compartida con la cuenta de servicio)")


def texto(respaldo):
    """Texto corto para la pantalla y para RegistroEventos."""
    base = f"{respaldo['nombre']} del {respaldo['fecha'].strftime('%d-%m-%Y')} ({respaldo['dias']} dias)"
    if respaldo["viejo"]:
        base += f" RESPALDO VIEJO: mas de {dias_aviso()} dias, actualizarlo"
    return base


def olvidar():
    """Solo para las pruebas: vuelve a leer la carpeta en la siguiente llamada."""
    _leidos.clear()


# ---------------------------------------------------------------- Uso directo
def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:       # noqa: BLE001
        pass
    if "--ver" not in sys.argv:
        print(__doc__)
        return
    donde = modo()
    print("=== Respaldo manual de fuentes (solo lectura) ===")
    if donde is None:
        print("No hay carpeta de respaldos: falta RESPALDO_FUENTES_FOLDER_ID en el .env "
              f"(o, con --local, la carpeta {carpeta_local()}).")
        return
    print("Carpeta: " + (f"COPIA LOCAL {carpeta_local()}" if donde == "local"
                         else "GOOGLE DRIVE (RESPALDO_FUENTES_FOLDER_ID)"))
    faltan = 0
    for fuente in PREFIJOS:
        try:
            r = leer(fuente)
        except RespaldoNoDisponible as e:
            faltan += 1
            print(f"  {fuente:<7} SIN RESPALDO: {e}")
            continue
        print(f"  {fuente:<7} {texto(r)} -> {r['resumen']}")
        for d in r["descartados"]:
            print(f"          (descartado, mas nuevo pero no sirve) {d}")
    print("LISTO: los dos respaldos se pueden usar." if not faltan
          else f"ATENCION: falta el respaldo de {faltan} fuente(s).")


if __name__ == "__main__":
    main()
