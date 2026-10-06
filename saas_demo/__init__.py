"""SaaS de demostracion de la panaderia (Sesion 2, 3-oct-2026).

Pantallas web para ver los pedidos digitales, marcar "Entregado" y anular (solo
administracion) sobre la COPIA de demostracion del Excel. Usa las funciones
compartidas del robot (pedidos_digitales.py), las mismas del bot de WhatsApp.

Se corre SIEMPRE desde la carpeta ROBOT PANADERO como paquete:
    py -m saas_demo.ensayo              (en la PC, copia temporal; no toca Drive)
    py -m saas_demo.probar_saas_demo    (prueba automatica, copia temporal)
En Render es un servicio aparte del bot (ver app.py).

|| 4-oct-2026: modulos por grupo, en el orden del Excel maestro (Regla #39):
    grupo1.py  Ventas despacho, descuento de dia de pago e Ingreso de inventario
    grupo2.py  Control de inventario (existencia minima, #63) y Hornadas
    grupo3.py  NODO INTEGRAL (celdas manuales, #64) y lectura de las dos tablas cruce maestras
|| 5-oct-2026: grupo3.py tambien trae 'Tabla combos dinamicos': configuracion de cada combo (#65).
|| 5-oct-2026 (tarde): y la composicion de cada combo (#85), en la misma pantalla.
|| 5-oct-2026 (noche): grupo4.py  Costo pan baguette y Costo cupcakes (lectura) y precio de referencia.
|| 5-oct-2026 (noche): grupo4.py tambien trae las proporciones de las hojas de costo, Planilla, Gasto operacional y Valor activos.
|| 6-oct-2026: grupo5.py  Flujo caja diario (caja chica del dia, #66) y Flujo caja anual (salidas de efectivo del mes, #67); solo administracion.
"""
