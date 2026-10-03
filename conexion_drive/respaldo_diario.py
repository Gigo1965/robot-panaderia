"""Respaldo diario del Excel maestro en Google Drive (fase de hosting, 2-oct-2026).

Que hace:
  Copia el contenido EXACTO del Excel maestro de Drive (byte por byte, sin editarlo ni
  recalcularlo) dentro de OTRO archivo de Drive, el archivo de respaldo (por ejemplo
  RESPALDO_DIARIO_MAESTRO.xlsx), que el usuario creo una sola vez y compartio como
  Editor con la cuenta de servicio del robot.
  La tarea diaria de GitHub Actions lo corre a las 7:00 a. m., ANTES de NODO INTEGRAL,
  precios y competencia. Como el respaldo se escribe una sola vez al dia, el historial
  de versiones de Drive de ESE archivo queda con una version por dia (el del maestro se
  llena rapido con cada pedido del bot).

Que NO hace: no toca el Excel maestro (solo lo lee), no crea archivos nuevos en Drive y
no borra nada.

Seguros (se detiene sin escribir si algo no coincide):
  - el archivo de respaldo NO puede ser el Excel maestro (mismo ID);
  - su nombre debe contener la palabra RESPALDO (asi un ID equivocado, por ejemplo el
    del cache de competencia, no se pisa por error);
  - debe ser un .xlsx real (no Google Sheets) y el robot debe poder editarlo;
  - despues de subir, la huella (md5) del respaldo debe ser igual a la del maestro.
  Si el respaldo ya es identico al maestro, no se sube nada (no gasta versiones).

De donde toma el ID del archivo de respaldo (sirve el ID solo o el enlace completo de
Drive "Copiar enlace"):
  - el argumento:        py -m conexion_drive.respaldo_diario "<ID o enlace>"
  - o la variable:       RESPALDO_MAESTRO_FILE_ID  (.env en la PC; Secret en GitHub)

Igual que el Excel maestro (Regla #12): el respaldo NUNCA se abre con Google Sheets.
Para recuperar un dia: Drive -> clic derecho en el respaldo -> Informacion del archivo ->
Administrar versiones -> descargar la version de la fecha buscada.
"""
import hashlib
import os
import re
import sys
from io import BytesIO

from . import config
from .drive_excel import (MIME_XLSX, ErrorConexionExcel, _explicar_error_google, _id_archivo,
                          _servicio_drive)
from .tiempo import ahora_local

PALABRA_OBLIGATORIA = "RESPALDO"
CAMPOS = "id,name,mimeType,md5Checksum,size,version,modifiedTime,capabilities(canEdit)"


def id_desde_texto(texto):
    """Acepta el ID solo o el enlace de Drive (.../d/<ID>/...). Devuelve el ID."""
    texto = (texto or "").strip().strip('"').strip("'")
    encontrado = re.search(r"/d/([A-Za-z0-9_-]{20,})", texto)
    return encontrado.group(1) if encontrado else texto


def _datos(servicio, id_archivo):
    return servicio.files().get(fileId=id_archivo, fields=CAMPOS, supportsAllDrives=True).execute()


def _descargar(servicio, id_archivo):
    from googleapiclient.http import MediaIoBaseDownload
    solicitud = servicio.files().get_media(fileId=id_archivo, supportsAllDrives=True)
    memoria = BytesIO()
    descargador = MediaIoBaseDownload(memoria, solicitud)
    terminado = False
    while not terminado:
        _, terminado = descargador.next_chunk()
    return memoria.getvalue()


def _subir(servicio, id_archivo, contenido):
    from googleapiclient.http import MediaIoBaseUpload
    media = MediaIoBaseUpload(BytesIO(contenido), mimetype=MIME_XLSX, resumable=True)
    return servicio.files().update(fileId=id_archivo, media_body=media, fields=CAMPOS,
                                   supportsAllDrives=True).execute()


def respaldar(id_respaldo):
    """Copia el Excel maestro al archivo de respaldo. Devuelve un dict con el resultado."""
    if config.ORIGEN_EXCEL == "local":
        raise ErrorConexionExcel("El respaldo diario trabaja solo con Google Drive (no con --local).")
    id_respaldo = id_desde_texto(id_respaldo)
    if not id_respaldo:
        raise ErrorConexionExcel(
            "Falta el ID del archivo de respaldo: pasalo como argumento o en RESPALDO_MAESTRO_FILE_ID.")
    id_maestro = _id_archivo()
    if id_respaldo == id_maestro:
        raise ErrorConexionExcel(
            "El ID del respaldo es el MISMO del Excel maestro. No se hizo nada.")

    servicio = _servicio_drive()
    try:
        respaldo = _datos(servicio, id_respaldo)
        maestro = _datos(servicio, id_maestro)
    except Exception as error:
        raise ErrorConexionExcel(
            _explicar_error_google(error).replace("EXCEL_MAESTRO_FILE_ID",
                                                  "ID del maestro o del respaldo")) from error

    nombre = respaldo.get("name") or ""
    if PALABRA_OBLIGATORIA not in nombre.upper():
        raise ErrorConexionExcel(
            f"El archivo de destino se llama '{nombre}' y no contiene la palabra "
            f"{PALABRA_OBLIGATORIA}: no parece el archivo de respaldo. No se hizo nada.")
    if respaldo.get("mimeType") != MIME_XLSX:
        raise ErrorConexionExcel(
            f"El respaldo '{nombre}' no es un .xlsx (es {respaldo.get('mimeType')}). "
            "No debe convertirse a Google Sheets. No se hizo nada.")
    if not (respaldo.get("capabilities") or {}).get("canEdit"):
        raise ErrorConexionExcel(
            f"El robot no puede editar '{nombre}'. Compartelo como Editor con "
            "robot-panaderia@robot-panaderia-canas.iam.gserviceaccount.com")

    try:
        contenido = _descargar(servicio, id_maestro)
    except Exception as error:
        raise ErrorConexionExcel(_explicar_error_google(error)) from error
    huella = hashlib.md5(contenido).hexdigest()
    resultado = {"maestro": maestro.get("name"), "respaldo": nombre, "bytes": len(contenido),
                 "huella": huella, "version_maestro": maestro.get("version"),
                 "momento": ahora_local()}
    if respaldo.get("md5Checksum") == huella:
        resultado["subido"] = False
        return resultado

    try:
        nuevo = _subir(servicio, id_respaldo, contenido)
    except Exception as error:
        raise ErrorConexionExcel(
            _explicar_error_google(error).replace("EXCEL_MAESTRO_FILE_ID",
                                                  "ID del respaldo")) from error
    if nuevo.get("md5Checksum") != huella:
        raise ErrorConexionExcel(
            "El respaldo se subio pero su huella NO coincide con la del maestro. Revisar en Drive.")
    resultado["subido"] = True
    resultado["version_respaldo"] = nuevo.get("version")
    return resultado


def main():
    argumentos = [a for a in sys.argv[1:] if not a.startswith("--")]
    destino = argumentos[0] if argumentos else os.getenv("RESPALDO_MAESTRO_FILE_ID")
    print("=== Respaldo diario del Excel maestro (Google Drive) ===")
    try:
        r = respaldar(destino)
    except ErrorConexionExcel as error:
        print(f"FALLO: {error}")
        sys.exit(1)
    print(f"Maestro : {r['maestro']} (version {r['version_maestro']}, {r['bytes']:,} bytes)")
    print(f"Respaldo: {r['respaldo']}")
    if r["subido"]:
        print(f"OK  Copiado el {r['momento']:%Y-%m-%d %H:%M} (hora CR). Huella igual a la del "
              f"maestro: {r['huella'][:8]}...")
    else:
        print("OK  El respaldo ya era identico al maestro: no se subio nada.")
    print("LISTO. El Excel maestro no se toco.")


if __name__ == "__main__":
    main()
