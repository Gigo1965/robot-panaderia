# Bot de WhatsApp — Panadería (versión 29-sep-2026, Sesión A del bot)

Bot **solo en español**, con **menús numerados**, conectado al Excel maestro a través
de las funciones compartidas del robot (las mismas que usará el SaaS, Regla #5).
Canal: **Sandbox de Twilio** (sin plantillas de Meta: el bot solo responde dentro de
la ventana de 24 horas que abre el cliente).

## Archivos

| Archivo | Qué hace |
|---|---|
| `conversation.py` | La conversación completa (saludo con el menú de productos → cantidades → quitar → envío o retiro → pago → resumen #28 → confirmar) y el comando `ENTREGADO WA-xxx`. No habla con Twilio ni con el Excel. |
| `conector_excel.py` | Conector con el Excel: catálogo de hoy, cotizaciones, registro y entregas, usando `catalogo.py` y `pedidos_digitales.py`. Traduce errores a mensajes para el cliente. |
| `fila_escrituras.py` | Las escrituras al Excel van una a la vez, en segundo plano. |
| `nucleo.py` | Une las tres piezas anteriores; lo usan igual el servidor y el simulador. |
| `app.py` | Servidor Flask para Twilio: revisa la firma, responde al instante y deja la escritura en la fila. |
| `simulador.py` | Chatear con el bot en la Terminal, sin Twilio, contra la copia local. |
| `config.py` | Configuración (lee el `.env` de la raíz). Sin claves ni teléfonos reales. |
| `setup_content_templates.py` | OBSOLETO: plantillas de la versión anterior. Se conserva como historia. |

## Reglas que aplica (decisiones del usuario, 29-sep-2026)

- Precios, combos, feriados/días de pago, SINPE y promoción del día se leen del Excel en el momento.
- El cliente ve el **total completo** antes de confirmar; si cambia antes del "Sí", se vuelve a cotizar.
- Forma de pago obligatoria (Efectivo / SINPE Móvil). Las líneas se quitan solo antes de confirmar.
- Sin opción "Hablar con alguien": el saludo muestra directamente el menú; para consultas o errores, el
  mensaje de pedido registrado y el recibo remiten a la panadería (atención humana = mejora para producción).
- Pedido sin confirmar: se borra solo tras **30 minutos sin mensajes** (el reloj se reinicia con cada mensaje).
- Se guarda el teléfono (#33 B) y el nombre de perfil; aviso en el mensaje de confirmación.
- `ENTREGADO WA-120` solo desde números autorizados (`BOT_NUMEROS_AUTORIZADOS` en el `.env`):
  marca la entrega (`marcar_entregado`) y envía el recibo al cliente.

## Probar sin Twilio (desde la carpeta ROBOT PANADERO)

```
py -m whatsapp_bot.simulador
```

Escribe en la copia local `datos_prueba\maestro_prueba.xlsx` cuando confirmas o entregas:
foto antes → simulación → restaurar la copia.

## Servidor para Twilio (fase de hosting)

```
py -m whatsapp_bot.app
```

En Render: `gunicorn --workers 1 --threads 4 whatsapp_bot.app:app` (un solo proceso, porque
las conversaciones viven en memoria). URL del webhook en el Sandbox: `https://<servidor>/webhook`.
