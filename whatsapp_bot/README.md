# Bot de WhatsApp — Panadería (Capa de conversación)

Este es el **primer piso** del bot: la conversación completa (saludo →
ubicación → combos/menú → armar pedido → confirmar), probada y
funcionando de forma aislada, **sin escribir todavía en el Excel
maestro**. Esa segunda parte se conecta después, cuando exista la
función Python compartida con la futura web del SaaS (ver
`Reglas_Globales_del_Proyecto`, Regla #4).

## Arquitectura de archivos

| Archivo | Qué hace |
|---|---|
| `conversation.py` | La máquina de estados de la conversación. Puro Python, sin Twilio — se puede probar sin credenciales. |
| `app.py` | El servidor Flask que conecta Twilio con `conversation.py`. |
| `setup_content_templates.py` | Script que se corre UNA vez para crear las plantillas de WhatsApp en Twilio (listas, botones). |
| `config.py` | Todas las credenciales y catálogo de productos, vía variables de entorno. |

## Por qué el cliente nunca escribe texto libre

Confirmado en el proyecto: el cliente arma su pedido tocando botones y
listas (no escribiendo códigos de producto). Esto se refleja en que
`conversation.py` solo usa `texto_libre` para el saludo inicial y para
la cantidad numérica — todo lo demás es `opcion_elegida_id` (el ID que
Twilio manda cuando el cliente toca un botón o una fila de una lista).

## Instalación

```bash
pip install -r requirements.txt
cp .env.example .env
# Edita .env con tus credenciales reales de Twilio
```

## Paso 1 — Activar el Sandbox de WhatsApp de Twilio

1. Entra a la Twilio Console → Messaging → Try it out → Send a WhatsApp message.
2. Sigue las instrucciones para unirte al Sandbox desde tu propio WhatsApp (escaneas un código QR o mandas un mensaje de "join ...").
3. Copia el número de Sandbox que te asignan a `TWILIO_WHATSAPP_FROM` en tu `.env`.

## Paso 2 — Crear las plantillas de contenido

```bash
python setup_content_templates.py
```

Copia los 4 `ContentSid` que imprime (empiezan con `HX...`) a tu `.env`.

## Paso 3 — Correr el servidor y exponerlo con ngrok

```bash
flask --app app run
```

En otra terminal:

```bash
ngrok http 5000
```

Copia la URL pública que te da ngrok (algo como `https://xxxx.ngrok-free.app`)
y configúrala en la Twilio Console, en el campo "WHEN A MESSAGE COMES IN"
del Sandbox, agregando `/webhook` al final (ej.
`https://xxxx.ngrok-free.app/webhook`).

## Paso 4 — Probar

Escríbele por WhatsApp al número del Sandbox. Deberías recibir el menú
principal y poder completar un pedido de prueba de principio a fin.

## Probar la lógica sin WhatsApp (recomendado primero)

`conversation.py` no depende de Twilio, así que se puede probar
directo en Python:

```python
from conversation import GestorConversaciones, procesar_mensaje_entrante

gestor = GestorConversaciones()
conv = gestor.obtener_o_crear("whatsapp:+50688887777")

r = procesar_mensaje_entrante(conv, texto_libre="Hola")
print(r.tipo_mensaje)  # "menu_principal_y_pedir_ubicacion"
```

## Pendientes conocidos (documentados a propósito, no son bugs)

- **`obtener_combos_activos_hoy()`** en `app.py`: hoy devuelve una lista
  vacía (placeholder). Falta conectarla a la lectura real de
  `'Tabla Combos Dinamicos'` vía Microsoft Graph API.
- **`escribir_pedido_en_excel()`** en `app.py`: lanza `NotImplementedError`
  a propósito. Debe reemplazarse por la función Python compartida con
  la web del SaaS quest se construya, para no duplicar la lógica de
  negocio en 2 lugares.
- **Lista de combos activos**: hoy se envía como texto plano en vez de
  una lista interactiva de Twilio, porque las plantillas de Twilio son
  estáticas y los combos cambian día a día. La solución recomendada es
  un trabajo diario (mismo patrón de "disparador por tiempo" que usa
  NODO INTEGRAL) que actualice una plantilla de contenido fija
  (`friendly_name="combos_del_dia"`) cada mañana — todavía no
  construido.
- **Persistencia de conversaciones**: `GestorConversaciones` guarda
  todo en memoria; si el servidor se reinicia, las conversaciones en
  curso se pierden. Aceptable para el volumen de una panadería, pero
  documentado por si se necesita cambiar a una base de datos simple
  más adelante.
