"""Grupo 3 - Integraciones externas.

3A (hecho): fuentes.py + nodo_integral.py -> clima, tipo de cambio, feriados y dias de pago
            en la hoja NODO INTEGRAL (Regla #4: NODO INTEGRAL es el UNICO que llama a las APIs).
3B (pendiente): distancia de Google Maps por pedido.
3C (pendiente): scraping de precios (Pali, Solera) y promocion de Musmanni.

|| 4-oct-2026: 3B y 3C ya estan hechos (google_maps.py, precios_proveedores.py, competencia.py).
   Nuevo: nodo_manual.py -> celdas MANUALES de NODO INTEGRAL (C3, C4, C8 y C22) para el SaaS
   (Decisiones Pendientes #64 y #80).
   Nuevo: activacion_promocion.py -> el robot escribe 'Activacion de Promocion' (C16): ON si hay
   combos activos o descuento general de dia de pago / feriado (#84). La tarea diaria la llama
   al final de nodo_integral.py; en la copia de demostracion, un boton del SaaS.
"""
