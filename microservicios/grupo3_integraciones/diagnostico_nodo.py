"""Diagnostico de las fuentes de NODO INTEGRAL: consulta internet y MUESTRA lo que encontro.
NO escribe nada en el Excel.

Uso (desde la carpeta ROBOT PANADERO):
    py -m microservicios.grupo3_integraciones.diagnostico_nodo
"""
from conexion_drive.tiempo import ahora_local

from . import fuentes

if __name__ == "__main__":
    hoy = ahora_local().date()
    print(f"Fecha en Costa Rica: {hoy}\n")

    print("1) Clima de Cañas (Open-Meteo)")
    try:
        print("   OK ", fuentes.obtener_clima())
    except fuentes.FuenteNoDisponible as error:
        print("   FALLO:", error)

    print("2) Tipo de cambio de referencia del BCCR (API de Hacienda)")
    try:
        print("   OK ", fuentes.obtener_tipo_cambio())
    except fuentes.FuenteNoDisponible as error:
        print("   FALLO:", error)

    print("3) Feriados de Costa Rica de pago obligatorio (libreria holidays)")
    try:
        print("   Hoy:", fuentes.nombre_feriado(hoy) or "no es feriado")
        import holidays
        try:
            lista = holidays.country_holidays("CR", years=hoy.year, language="es")
        except Exception:
            lista = holidays.country_holidays("CR", years=hoy.year)
        for fecha, nombre in sorted(lista.items()):
            print(f"      {fecha}  {nombre}")
    except fuentes.FuenteNoDisponible as error:
        print("   FALLO:", error)

    print("4) Dia de pago (15 y 30; febrero: ultimo dia)")
    print("   Hoy:", "SI" if fuentes.es_dia_de_pago(hoy) else "NO")
