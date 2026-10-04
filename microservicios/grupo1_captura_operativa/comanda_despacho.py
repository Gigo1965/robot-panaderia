"""Microservicio: comanda de despacho = varias ventas de mostrador en UN solo guardado (#70).

Decision del usuario (3-oct-2026, opcion B): cuando un cliente lleva varios productos, el
despachador arma la comanda en el SaaS, VE EL TOTAL A COBRAR antes de confirmar y, al
confirmar, el robot escribe UNA FILA POR PRODUCTO en 'Ventas despacho' (Tabla6), todas con
la misma fecha y hora. La estructura del Excel NO cambia: la hoja sigue teniendo una fila
por producto, como dice su mapa de celdas; no hay numero de comanda en la hoja.

No es una regla nueva de precios: usa EXACTAMENTE las de ventas_despacho.py
(registrar_venta_despacho), de donde importa sus piezas:
  - precio en 2 pasos: combo primero (solo ACTIVOS), TABLA CRUCE MAESTRA despues;
  - feriado o dia de pago: no hay combos y el precio unitario ya lleva el descuento
    general de 'Interface digital'!B5;
  - valores FIJOS (fecha, hora, producto, cantidad, precio); Total = formula de la fila;
  - duplicado = misma Fecha + Hora + Producto + Cantidad (no se escribe dos veces).

Ademas, igual que en los pedidos digitales (el cliente ve el total completo antes de
confirmar): si el total cambio entre lo que vio el despachador y el momento de registrar
(p. ej. se apago un combo o cambio un precio), se rechaza sin escribir nada.
La comanda entra completa o no entra. Dos lineas del mismo producto se suman en una.
Queda UN evento VENTA_DESPACHO por comanda en RegistroEventos, con el total cobrado.

    cotizar_comanda_despacho([{'producto': 'P01', 'cantidad': 3}, {'producto': 'L01', 'cantidad': 1}])
    registrar_comanda_despacho(lineas, momento, total_cotizado=3630.57)
"""
from conexion_drive import editar_excel, leer_excel
from conexion_drive.tablas import agregar_fila, obtener_tabla, registrar_evento

from .catalogo import Catalogo
from .ventas_despacho import (
    FORMULA_TOTAL, HOJA, ORDEN_CORRECTA, TABLA, _buscar_venta, _separar_fecha_hora, _validar_cantidad,
)

MAXIMO_LINEAS = 12


class ComandaInvalida(ValueError):
    """La comanda esta vacia o tiene demasiadas lineas. No se escribio nada."""


class TotalCambio(Exception):
    """El total ya no es el que vio el despachador. No se escribio nada."""


def _calcular(lineas, catalogo):
    """Resuelve y suma. Lanza el error del catalogo si un producto no se puede vender hoy."""
    if not lineas:
        raise ComandaInvalida("La comanda esta vacia. No se escribio nada.")
    reglas = catalogo.reglas
    pct = reglas["pct_general"] if reglas["descuento_general"] else 0.0
    juntas = {}
    for linea in lineas:
        cantidad = _validar_cantidad(linea.get("cantidad"))
        item = catalogo.resolver(linea.get("producto"))               # combos inactivos: rechazados aqui
        precio = round(float(item["precio"]), 2)
        if pct:                                                       # misma regla que registrar_venta_despacho
            precio = round(precio * (1 - pct), 2)
        if item["nombre"] in juntas:
            juntas[item["nombre"]]["cantidad"] += cantidad
        else:
            juntas[item["nombre"]] = {"producto": item["nombre"], "codigo": item["codigo"], "origen": item["origen"],
                                      "cantidad": cantidad, "precio": precio}
    if len(juntas) > MAXIMO_LINEAS:
        raise ComandaInvalida(f"La comanda tiene mas de {MAXIMO_LINEAS} productos distintos. No se escribio nada.")
    resultado = list(juntas.values())
    for linea in resultado:
        linea["total"] = round(linea["precio"] * linea["cantidad"], 2)
    nota = f"{pct:.0%} {reglas['detonador_general']}" if pct else ""
    return resultado, round(sum(l["total"] for l in resultado), 2), nota


def cotizar_comanda_despacho(lineas, wb_valores=None):
    """Lo que se le muestra al despachador ANTES de confirmar. No escribe nada."""
    calculadas, total, nota = _calcular(lineas, Catalogo(wb_valores or leer_excel()))
    return {"lineas": calculadas, "total": total, "descuento": nota}


def registrar_comanda_despacho(lineas, momento=None, total_cotizado=None, origen="SaaS"):
    """Registra la comanda: una fila por producto en Tabla6, en un solo guardado.

    lineas:         [{'producto': codigo o nombre exacto, 'cantidad': entero > 0}, ...]
    momento:        fecha y hora de la venta (el SaaS envia siempre el momento original).
    total_cotizado: el total que vio el despachador; si ya no coincide, no se escribe.
    """
    fecha, hora = _separar_fecha_hora(momento)

    def _editar(wb, wb_valores):
        calculadas, total, nota = _calcular(lineas, Catalogo(wb_valores))
        if total_cotizado is not None and round(float(total_cotizado), 2) != total:
            raise TotalCambio(
                f"El total cambio: se mostro {float(total_cotizado):,.2f} y ahora es {total:,.2f} "
                "(cambio un precio o un combo). No se escribio nada; revise la comanda y confirme de nuevo.")
        ws, tabla = obtener_tabla(wb, HOJA, TABLA)

        registradas, duplicadas = [], []
        for linea in calculadas:
            existente = _buscar_venta(ws, tabla, fecha, hora, linea["producto"], linea["cantidad"])
            if existente:
                duplicadas.append({**linea, "fila": existente["_fila"]})
                continue
            fila = agregar_fila(ws, tabla, {
                "Fecha": fecha, "Hora": hora, "Producto": linea["producto"], "Cantidad": linea["cantidad"],
                "Precio": linea["precio"], "Total": FORMULA_TOTAL, "Estado de orden": ORDEN_CORRECTA,
            }, columnas_clave=["Fecha", "Producto", "Cantidad"])
            registradas.append({**linea, "fila": fila})

        base = {"lineas": registradas, "duplicadas": duplicadas, "total": total,
                "fecha": fecha.date().isoformat(), "hora": hora.isoformat()}
        if not registradas:
            return {"hubo_cambios": False, "estado": "DUPLICADO", **base}

        filas = ", ".join(str(l["fila"]) for l in registradas)
        registrar_evento(
            wb, modulo="Ventas despacho", tipo_evento="VENTA_DESPACHO", entidad="Tabla6",
            id_entidad=f"filas {filas}" if len(registradas) > 1 else f"fila {filas}", accion="INSERT",
            estado=ORDEN_CORRECTA, origen=origen, valor_referencia=round(sum(l["total"] for l in registradas), 2),
            detalle="comanda de despacho: " + "; ".join(
                f"{l['cantidad']} x {l['producto']} a {l['precio']} ({l['codigo']}, {l['origen']})" for l in registradas)
            + (f" | {nota}" if nota else "") + f" | total a cobrar {total:,.2f}",
        )
        return {"hubo_cambios": True, "estado": "REGISTRADA", **base}

    return editar_excel(_editar, descripcion=f"Comanda de despacho ({len(lineas)} linea(s))")
