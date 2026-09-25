# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V5

*Última actualización: 24-sep-2026 (cierre de sesión, noche). Reemplaza a la V4. Guarda este archivo en la
carpeta del proyecto (junto a CLAUDE.md) y súbelo al inicio de cualquier chat nuevo con Claude.*

**Nota para quien retome este proyecto (incluido Claude en un chat nuevo):**
el usuario (Francisco) es principiante total en Git, terminal y programación.
Necesita instrucciones muy concretas y paso a paso, con nombres exactos de
botones y menús. Los comandos de Terminal (escribir texto y dar Enter) le
funcionan mejor que buscar iconitos o arrastrar con el mouse (elimina error
humano). Claude actúa como **tutor**: explicar qué se construye y por qué.
**Protocolo del proyecto**: anunciar lo que se va a hacer y esperar confirmación;
nunca perder información; diff celda por celda antes/después; validar con
LibreOffice; nunca escribir texto ilustrativo que empiece con "=".
**Este archivo está en el repo (va a GitHub): NUNCA poner claves aquí.**

---

## 1. Qué es el proyecto

Estudio de factibilidad de una panadería (Cañas, Guanacaste, Costa Rica) con
un Excel maestro (financiero, 25 hojas) + robot en Python que lo mantiene
actualizado desde un SaaS y un bot de WhatsApp. Proyecto final de maestría
(tutores en México). Todo en la nube y automático. **Los tutores solo ven el
código Python y todo funcionando; nunca el libro de mapas (es solo del usuario).**

## 2. Arquitectura confirmada (no cambiar sin razón)

1. **Excel**: Google Drive + Cuenta de Servicio (sin login humano).
   - `robot-panaderia@robot-panaderia-canas.iam.gserviceaccount.com`, proyecto `robot-panaderia-canas`
   - `EXCEL_MAESTRO_FILE_ID=1uOY84vrtjmV_YvReHWxcHlumEATpCnlw`
   - descargar → editar con openpyxl → recalcular con LibreOffice → subir.
     Si hay choque de versión, `editar_excel` repite la edición (hasta 3 veces).
   - Compartir > engranaje: "Editors" en "People who can download, copy, and print".
2. **Bot de WhatsApp**: Twilio, selección guiada. **Bilingüe** (es/en) en la fase del bot.
3. **Hosting planeado**: Render (bot) + GitHub Actions (tareas programadas).
4. **NODO INTEGRAL** es el único proceso que llama APIs por horario (clima,
   feriados, tipo de cambio). Excepción: Google Maps por pedido (Regla #14).
5. Sin "cliente único": todo en unidades y pedidos.
6. `insert_rows()` no ajusta fórmulas: el robot escribe cada fórmula.
7. Power BI después de datos reales, vía Drive para Escritorio (nunca enlace público).
8. **El Excel NUNCA se abre con Google Sheets** (rompe `_xlfn.TEXTJOIN`).
   El escudo `_xlfn` v2 lo repara y ahora **recalcula** al reparar.
9. Tipo de cambio (NODO C7) = referencia BCCR vía API de Hacienda.
10. Escrituras en tiempo real (~250/día) por la MISMA función compartida.
11. Mapas de celdas = base del código; todo cambio se refleja en el libro de mapas.
12. **Cambio de producto o precio → se cambia en TODOS los puntos que afecta**:
    todas las hojas del Excel (TABLA CRUCE MAESTRA/MP, Ingreso de Inventario,
    Control de Inventario, Ventas despacho, combos...) y el libro de mapas.
    Los nombres de producto funcionan como LLAVE (Control de Inventario usa SUMIF por nombre).
13. **Claves**: el código las lee SIEMPRE del `.env`. El libro de mapas guarda
    toda la info de APIs (URL, clave, etc.) por decisión del usuario (Regla #15).

## 3. Reglas de negocio

- Captura SaaS con autocompletado; el robot valida códigos.
- Despacho y venta en línea: ingresos separados.
- Combos iguales en ambos canales; solo ACTIVOS; ya traen descuento.
- Días de pago 15 y 30 (feb: último día; fin de semana cuenta) → 5%, sin combos.
- Feriados: solo los 9 obligatorios → 5%, sin combos.
- Hornadas suman producción y rebajan materia prima; anular = G2 en 0 + fila negativa.
- Ingresos con precio de referencia; costos dinámicos.
- **Pedidos del bot**: el cliente ve el TOTAL COMPLETO (productos − descuentos +
  envío) antes de confirmar. Envío = km × Tarifa (NODO C3), sin IVA, redondeado
  hacia arriba a ₡50; radio 10 km; retiro en local = envío 0.
- **Reemplazar materia prima que ya no existe**: por PRECIO y FUNCIONALIDAD,
  siempre preguntando al usuario; si la calidad es semejante gana el mejor precio.
  Levadura: la más barata (seca o fresca solo cambia el proceso, según el usuario).
- **Leche**: siempre Dos Pinos semidescremada 1 L; respaldo Sabemas UHT 1 L si no hay.

## 4. Orden de construcción

| Grupo | Estado |
|---|---|
| 1 — Captura operativa | ✅ |
| 2 — Inventario | ✅ |
| 3 — Integraciones | ✅ 3A · ✅ 3B (partes 1 y 2) · 🔶 3C (decisiones listas, falta m003 + código) |
| 4 — Costeo | ⏳ |
| 5 — Financiero | ⏳ |

## 5. Carpeta en la PC

`C:\Users\franf\Desktop\ROBOT PANADERO` (comillas en comandos). PowerShell en VS Code.
Descargas: a `Escritorio\download` (nunca Descargas); de ahí con `Copy-Item`.
**Respaldo del Excel antes de la m003**: `Escritorio\respaldo_antes_m003`
(Excel descargado de Drive el 24-sep 17:45, 219 KB, nombre
`Estudio_Financiero_Panadero_CON_ROUND_v68_DRIVE_CALCULADO.xlsx` — el nombre
dice v68 pero el contenido es el actual de Drive). No se toca nunca.

## 6. Entorno y comandos

- Python 3.14; usar `py -m pip ...` y `py -m <modulo>`. LibreOffice instalado.
- Probar conexión: `py -m conexion_drive.probar_conexion` (25 hojas)
- Pedidos (local): `py -m microservicios.grupo1_captura_operativa.probar_pedidos_digitales`
- Google Maps: `py -m microservicios.grupo3_integraciones.google_maps <lat> <lng>`
- Romper copia de prueba (test escudo): `py romper_prueba.py`
- Recalcular copia local a mano:
  `py -c "from pathlib import Path; from conexion_drive.recalculo import reparar_prefijos_xlfn, recalcular_bytes; p=Path('datos_prueba/maestro_prueba.xlsx'); c,r=reparar_prefijos_xlfn(p.read_bytes()); p.write_bytes(recalcular_bytes(c)); print('OK')"`
- Probar clave Gemini (debe dar 200):
  `py -c "import os,http.client as h; from conexion_drive import config; c=h.HTTPSConnection('generativelanguage.googleapis.com',timeout=20); c.request('GET','/v1beta/models',headers={'x-goog-api-key': os.getenv('GEMINI_API_KEY')}); r=c.getresponse(); print(r.status, r.read()[:200].decode())"`
- Verificar que un archivo copiado es el nuevo: `Select-String -Path "<archivo>" -Pattern "<marca>"`

## 7. Git / GitHub

- Repo privado https://github.com/Gigo1965/robot-panaderia.git (main).
- Commits: V4, cd7db42 (G2), c52e289 (NODO v3), 67d3361 (escudo), 5becbcc,
  a3a69f1 (3B parte 1), **e0c529d (3B parte 2: envío Google Maps, cotizar → registrar)**,
  **20d2329 (escudo _xlfn v2: recalcula tras reparar)**.
- `git status` → `git add .` → `git commit -m "..."` → `git push`.
- `git grep AIza` debe salir vacío (ninguna clave en el repo).
- Nunca editar en la web de GitHub.

## 8. Documentos y claves (fuera del repo)

- **Libro de mapas**: vigente **V9** (clave de Maps restaurada, Reglas #14 y #15,
  pestaña "Cambios V9"). Próximo: **V10** con la m003 y el 3C.
- **Claves en el `.env`** (valores en el `.env` y en el libro de mapas, NO aquí):
  - `GOOGLE_MAPS_API_KEY`
  - `GEMINI_API_KEY` — proyecto "Default Gemini Project" (nivel gratuito,
    `projects/253837609158`). Formato nuevo de Google: empieza con `AQ.` (válido).
    La clave del proyecto con prepago (552251224672) daba 401: en desuso.
- Drive: "Conservar para siempre" no funcionó desde la cuenta del usuario
  (probablemente porque la versión la subió el robot) → respaldo manual a la PC.

## 9. Dónde quedamos (24-sep-2026, cierre)

**Hecho hoy:**
- 3B parte 2 completa y probada con Google real (Cañas 1.05 km, 4 min, ₡550;
  Liberia 48.77 km fuera de zona). Funciones: `cotizar_envio`, `cotizar_pedido`,
  `registrar_pedido_digital(..., distancia_km, minutos_entrega, total_cotizado)`.
  NODO C9/C10 = minutos y km del último pedido.
- Escudo `_xlfn` v2 probado (copia rota a propósito → reparó, recalculó, pedidos OK).
- Libro de mapas V9.
- Clave Gemini funcionando (Gemini 2.5 Flash).
- 3C: investigación completa de proveedores (ver abajo).

**3C — Método decidido:** capa principal = **API pública de catálogo de cada
tienda** (gratis, sin IA, no depende del diseño): MaxiPalí (VTEX)
`https://www.maxipali.co.cr/api/catalog_system/pub/products/search?fq=alternateIds_Ean:<EAN>`
(precio en `items[0].sellers[0].commertialOffer.Price`); Solera (WooCommerce)
`https://solera.co.cr/wp-json/wc/store/v1/products?slug=<slug>` (precio en
`prices.price`, `currency_minor_unit` 0). Respaldo = ScrapeGraphAI con Gemini.
Guardar SKU/EAN para buscar si cambia la URL. Frecuencia: precios 1 vez al día;
Musmanni 1–2 veces al día. Parámetros del mapa: timeout 30 s, 4 reintentos,
pausa 2.5–5.5 s, secuencial. Si falla: NO tocar la celda. Límite de cambio de
precio >50% = no escribir y marcar (se activa DESPUÉS de la primera corrida real).

**MaxiPalí — TABLA CRUCE MAESTRA col D (verificados 24-sep):**

| Fila | Producto | EAN | Precio |
|---|---|---|---|
| 10 | Leche Semidescremada Dos Pinos -1 L (respaldo: Sabemas UHT 1 L, EAN 7441078217694, ₡980) | 7441001601064 | ₡1,050 |
| 11 | Queso Fresco Dos Pinos Semiduro Zarcero - 300g | 7441013503189 | ₡2,300 |
| 12 | Margarina Numar caja de 4 Uds - 500 g | 0751963110014 | ₡1,000 |
| 13 | Gaseosa Coca Cola Regular - 600 ml | 7441003500488 | ₡750 |
| 14 | Natilla Dos Pinos La Ganja Con Sal - 300g | 7441001608728 | ₡840 |

**Solera — TABLA CRUCE MAESTRA MP (catálogo completo 291 productos; 8 de 11 URLs viejas ya no existen):**

| Fila | Cód | Producto nuevo (nombre exacto Solera) | SKU | Precio | Col. E |
|---|---|---|---|---|---|
| 10 | Me01 | MEJORANTE PANNIVITA PRO-PAN 1 KG HOMO-PLUS | 0549 | 1815 | D10/1000 |
| 11 | Ha01 | HARINA FLORES 25 KG FUERTE | 0016 | 17510 | D11/25000 |
| 12 | Az01 | AZUCAR DOÑA MARIA 2KG (24 KG) (slug nuevo) | 0021 | 16490 | D12/24000 |
| 13 | Sa01 | SAL SARU COMUN ECONOMICA 25 KG SACO BRIN | 0995 | 3310 | D13/25000 |
| 14 | Ma02 | MANTECA MAX PAN BULK 25 KGS | 0030 | 28845 | D14/25000 |
| 15 | Le01 | LEVADURA FRESCA GOLONDRINA 400 GR CIP | 0137 | 735 | D15/400 |
| 16 | Me02 | MERMELADA GUAYABA HORNEO 5 KG CCR CAJA C2 (FRUTA TRITURADA) | 0957 | 4140 | D16/5000 |
| 17 | Bo01 | BOLSAS PAPEL KRAFT 2 BAGUETTE 100 UNDS BIODEGRADABLE | 0516 | 2400 | D17/100 |
| 18 | Bu01 | BURBUJA FOAM #6 125 UNID POSTRE BL | 0179 | 7275 | D18/125 |
| 19 | Ca01 | Café Naranjo puro 250gr x 10 Unidades | 0474 | 20755 | D19/2500 |
| 20 | Ma03 | MARGARINA TUMBA NUMAR 7.5 KILOS 60 BARRITAS X125GR (slug nuevo) | 0034 | 15390 | D20/7500 |

URL de cada una: `https://solera.co.cr/producto/<slug>/` (slugs en el catálogo;
p. ej. `mejorante-pannivita-pro-pan-1-kg-homo-plus`, `azucar-dona-maria-2kg-24-kg`,
`sal-saru-comun-economica-25-kg-saco-brin`, `levadura-fresca-golondrina-400-gr-cip`,
`mermelada-guayaba-horneo-5-kg-ccr-caja-c2-fruta-triturada`,
`bolsas-papel-kraft-2-baguette-100-unds-biodegradable`, `burbuja-foam-6-125-unid-postre-bl`,
`cafe-naranjo-puro-250gr-x-10-unidades`, `margarina-tumba-numar-7-5-kilos-60-barritas-x125gr`,
`harina-flores-25-kg-fuerte`, `manteca-max-pan-bulk-25-kgs`).
Ojo: "SAL ROCA 25 KG ... NO ES PARA CONSUMO HUMANO" existe en Solera: nunca usarla.

**Musmanni** (NODO C11): `https://musmanni.com/promociones-2/` — pendiente captura
+ código fuente; lectura del banner con **Gemini** (el mapa decía Claude: anotar "|| V10").

**Siguientes pasos (en este orden):**
1. **Migración m003** (chat nuevo): el usuario sube el Excel ACTUAL de Drive
   (el del respaldo). Claude aplica m003 como script: renombrar cada producto
   en TODAS las hojas (Suli → "Leche Semidescremada Dos Pinos -1 L" en 11 celdas:
   TCM C10, Control de Inventario A7, Ingreso de Inventario C4/C7, Ventas despacho
   E22/E42/E44/E147 —estas decían "grasa 1 L" sin guion y no se descontaban—,
   Tabla Combos D5/K5, DiccionarioDatos N80; y los productos de Solera en MP col C,
   Ingreso de Inventario col C, Control de Inventario col A) + conversiones col E.
   Diff celda por celda + LibreOffice + reporte + **libro de mapas V10** (incluye
   clave Gemini y lista de proveedores).
2. El usuario sube el Excel a Drive como **NUEVA VERSIÓN del mismo archivo**:
   clic derecho → Información del archivo → Gestionar versiones → **Subir nueva
   versión** (NUNCA "Subir archivo": cambia el ID y el robot pierde el acceso).
   Luego `py -m conexion_drive.probar_conexion`.
3. Código `grupo3_integraciones/precios_proveedores.py` (API + respaldo Gemini),
   prueba local, commit.
4. Musmanni (`musmanni.py` con Gemini).
5. Grupos 4 y 5. 6. Bot de WhatsApp real (bilingüe, función compartida).

**Pendientes para el Grupo 5:**
- Anular una línea de pedido digital: el registro original se mantiene y se
  agrega una **fila negativa** (asiento de reversión). Definir estado, fecha y envío.
- "Ventas digitales" de Flujo caja suma TOTAL_A_COBRAR, que incluye el envío.

**Pendientes de hosting:** minutos de Actions (ping a Render); respaldo diario
del Excel ("Keep forever" vía Drive API, límites de versiones).

## 10. Para continuar en un chat nuevo

Subir: **este archivo (V5)**, el **Excel actual de Drive** (el de
`respaldo_antes_m003`, sin abrirlo) y el **libro de mapas V9**. Decir:
"Continuemos con la migración m003 del Grupo 3C".
