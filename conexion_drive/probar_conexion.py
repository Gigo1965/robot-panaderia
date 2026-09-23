"""Prueba de la conexion con el Excel maestro en Google Drive.

Uso (desde la carpeta ROBOT_PANADERIA, en la Terminal):

    python -m conexion_drive.probar_conexion              <- solo LEE (seguro)
    python -m conexion_drive.probar_conexion --escribir   <- agrega 1 evento de prueba
                                                              en la hoja RegistroEventos
"""
import logging
import sys
import time

from . import config
from .drive_excel import editar_excel, leer_excel, obtener_metadatos
from .recalculo import encontrar_soffice
from .tablas import leer_tabla, obtener_tabla, registrar_evento

logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")


def prueba_lectura():
    print("\n=== 1) Conectando con Google Drive ===")
    print(f"Modo: {config.ORIGEN_EXCEL}")
    datos = obtener_metadatos()
    print(f"OK  Archivo: {datos.get('name')}")
    print(f"    Version en Drive: {datos.get('version')}   Modificado: {datos.get('modifiedTime', '-')}")

    print("\n=== 2) Descargando y abriendo con openpyxl ===")
    inicio = time.monotonic()
    wb = leer_excel(solo_valores=True)
    print(f"OK  Descargado en {time.monotonic() - inicio:.1f} s")
    print(f"    Hojas encontradas: {len(wb.sheetnames)} (se esperan 24)")

    print("\n=== 3) Leyendo precios de TABLA CRUCE MAESTRA ===")
    ws, tabla = obtener_tabla(wb, "TABLA CRUCE MAESTRA", "TablaProductos")
    for producto in leer_tabla(ws, tabla):
        print(f"    {producto['Codigo']:<6} {producto['Nombre tecnico completo'][:45]:<45} "
              f"Precio: {producto['Precio']}")

    print("\n=== 4) Buscando LibreOffice (necesario para escribir) ===")
    print(f"OK  {encontrar_soffice()}")


def prueba_escritura():
    print("\n=== 5) Prueba de ESCRITURA: 1 evento en RegistroEventos ===")

    def agregar_evento(wb, wb_valores):
        id_evento = registrar_evento(
            wb, modulo="Sistema", tipo_evento="PRUEBA_CONEXION", entidad="Libro Maestro",
            id_entidad="conexion_drive", accion="Prueba de escritura del robot", estado="OK",
            detalle="Evento creado por probar_conexion.py; puede eliminarse.",
        )
        return {"hubo_cambios": True, "id_evento": id_evento}

    inicio = time.monotonic()
    resultado = editar_excel(agregar_evento, descripcion="Prueba de conexion")
    print(f"OK  {resultado['id_evento']} guardado en {time.monotonic() - inicio:.1f} s")
    print("    Abre el Excel en Google Drive -> hoja RegistroEventos para verlo.")


if __name__ == "__main__":
    prueba_lectura()
    if "--escribir" in sys.argv:
        prueba_escritura()
    print("\nTODO BIEN.")
