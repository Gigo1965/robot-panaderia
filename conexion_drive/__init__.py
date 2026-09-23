"""Conexion del robot con el Excel maestro en Google Drive.

Patron unico del proyecto (Regla #1 del libro de mapas):
    descargar -> editar en memoria con openpyxl -> recalcular -> subir

Uso tipico desde un microservicio:

    from conexion_drive import editar_excel

    def mi_cambio(wb, wb_valores):
        # wb         = libro con FORMULAS (este es el que se edita)
        # wb_valores = el mismo libro con VALORES ya calculados (solo para leer)
        ...
        return {"hubo_cambios": True}

    editar_excel(mi_cambio, descripcion="Registrar venta")
"""
from .drive_excel import (
    ConflictoDeVersion,
    ErrorConexionExcel,
    editar_excel,
    leer_excel,
    obtener_metadatos,
)

__all__ = [
    "ConflictoDeVersion",
    "ErrorConexionExcel",
    "editar_excel",
    "leer_excel",
    "obtener_metadatos",
]
