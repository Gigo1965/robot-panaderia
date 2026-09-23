# Proyecto: Robot para Panadería (Estudio de Factibilidad)

## Qué es esto
Sistema para una panadería en Costa Rica: un Excel maestro (estudio financiero
completo, 26 hojas) + un robot en Python que lo mantiene actualizado a partir
de un SaaS y un bot de WhatsApp. El Excel YA está construido y validado
financieramente (VAN y TIR positivos). Lo que falta es el robot.

## Arquitectura — reglas que NUNCA se rompen

1. **Conexión a Excel**: el robot usa **Google Drive + una Cuenta de Servicio**
   (Service Account) — 100% automático, sin ningún login humano, nunca. El
   robot descarga el archivo `.xlsx` completo, lo edita EN MEMORIA con
   `openpyxl`, y sube la versión actualizada de vuelta a Google Drive (patrón
   "descargar → editar → subir"). El archivo se mantiene como `.xlsx` real,
   SIN convertir a Google Sheets. (CORREGIDO 17-sep-2026: se descartó
   Microsoft Graph API por restricciones de licencias — ver
   `Reglas_Globales_del_Proyecto`, Regla #1, para el detalle completo).
   IMPORTANTE: `openpyxl.insert_rows()` NO ajusta el texto de fórmulas
   existentes al insertar filas — cualquier fórmula debe escribirse
   explícitamente en cada celda nueva, nunca asumir que se "corre sola".
2. **Fuente única de verdad**: NODO INTEGRAL es el ÚNICO proceso que llama
   APIs externas (clima, feriados, tipo de cambio, Google Maps), por
   disparadores de TIEMPO (cron), no por evento de usuario. El SaaS y el Bot
   de WhatsApp SOLO LEEN ese estado ya actualizado — nunca llaman las APIs
   externas por su cuenta.
3. **Bot de WhatsApp = Twilio** (no Meta Cloud API directo, no 360dialog).
   El cliente NUNCA escribe texto libre para armar un pedido — todo es por
   selección guiada (Quick Reply / List Picker, plantillas de contenido de
   Twilio creadas una vez con `setup_content_templates.py`).
4. **Una sola función de escritura a Excel, compartida** entre el bot de
   WhatsApp y la futura web del SaaS — nunca duplicar la lógica de negocio
   (cálculo de subtotal, descuentos, envío) en dos lugares.
5. **Regla de redondeo global**: todos los montos monetarios del proyecto,
   en cualquier hoja, se manejan con exactamente 2 decimales.
6. **Historial vs. configuración**: hojas de eventos (Ventas despacho,
   Historico Diario, Interface digital, Detalle comanda digital) el robot
   SOLO HACE INSERT con valores fijos, nunca fórmula viva, nunca UPDATE
   salvo anulación explícita de una línea puntual. Hojas de configuración
   (planilla, gasto operacional, Punto de Equilibrio) SÍ permiten UPDATE.
7. **El concepto de "cliente único"/"perfil de cliente" NO existe** en el
   sistema — todo se maneja en unidades y pedidos, nunca pidas ni modeles
   un dato de "cantidad de clientes".

## Orden de construcción (5 grupos de microservicios)

Construir en este orden — cada grupo depende de que el anterior ya funcione:

1. **Captura operativa** (más urgente): Ventas despacho → Interface digital
   → Detalle comanda digital → Ingreso de Inventario
2. **Inventario / agregación derivada**: Control de Inventario → Historico
   Diario
3. **Integraciones externas**: NODO INTEGRAL → TABLA CRUCE MAESTRA → TABLA
   CRUCE MAESTRA MP → Tabla Combos Dinamicos
4. **Costeo y configuración**: costo pan baguette → costo cupcakes →
   planilla → gasto operacional → valor activos
5. **Financiero / análisis** (el más complejo, depende de todo lo anterior):
   Flujo caja diario → flujo caja anual → flujo caja proyecto → Punto de
   Equilibrio → TIR VAR → proyeccion ventas → proyeccion anual ventas

Power BI se conecta después de que los Grupos 1-2 tengan datos reales
(nunca antes — mostraría dashboards vacíos).

## Estilo de código
- Lógica **humana, no sobre-optimizada** — preferir claridad sobre
  elegancia clever.
- Estructura de **mini-servicios**, un archivo/módulo por responsabilidad.
- Comentarios y nombres de variables de dominio en **español** (materia
  del negocio: "pedido", "hornada", "combo", "cantidad_producida").
- Cada módulo nuevo se prueba de forma aislada antes de conectarlo (ver
  `whatsapp_bot/conversation.py` como ejemplo — se probó sin Twilio antes
  de integrarlo a `app.py`).

## Ya construido (no reconstruir desde cero)
- `whatsapp_bot/` — capa de conversación del bot (saludo → ubicación →
  combos/menú → armar pedido → confirmar), probada. NO escribe a Excel
  todavía (`escribir_pedido_en_excel()` es un placeholder intencional,
  ver Regla #4 de arquitectura arriba).

## Fuente de verdad para el detalle celda por celda
El "libro de mapas de celdas" (fuera de este repo, en el Excel de mapas)
tiene, para cada una de las 25 hojas: mapeo celda por celda, quién
escribe cada dato (usuario vía SaaS/Bot, o fórmula automática), el orden
de cálculo, y notas de diseño del robot. Consultar ese libro antes de
escribir el código de cada microservicio — no asumir la estructura de una
hoja sin confirmarla ahí primero.
