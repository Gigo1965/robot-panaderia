# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V7

*Última actualización: 25-sep-2026 (cierre de sesión, ~01:00). Reemplaza a la V6. Guarda este archivo en la
carpeta del proyecto (junto a CLAUDE.md) y súbelo al inicio de cualquier chat nuevo con Claude.*

**Nota para quien retome este proyecto (incluido Claude en un chat nuevo):**
el usuario (Francisco) es principiante total en Git, terminal y programación.
Necesita instrucciones muy concretas y paso a paso, con nombres exactos de
botones y menús. Los comandos de Terminal (escribir texto y dar Enter) le
funcionan mejor que buscar iconitos o arrastrar con el mouse (elimina error
humano). Claude actúa como **tutor**: explicar qué se construye y por qué.
**Protocolo del proyecto**: anunciar lo que se va a hacer y esperar confirmación;
nunca perder información; diff celda por celda antes/después (doble revisión:
el script revisa antes/después + un script independiente compara TODO el libro);
validar con LibreOffice; nunca escribir texto ilustrativo que empiece con "=".
Nunca se borran archivos de la PC sin pedirlo y explicarlo.
**Vocabulario (precisión para la tesis)**: Solera y MaxiPalí son PROVEEDORES; cada
uno tiene una API DE CATÁLOGO (fuente de datos JSON); se consultan 16 PRODUCTOS y
se escriben en 2 TABLAS del Excel. Nunca decir "tiendas". Musmanni y Musi son COMPETIDORES
(ambos de FIFCO); cada uno publica promos en un sitio WordPress.
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
| 3 — Integraciones | ✅ 3A · ✅ 3B · 🔶 3C (✅ m003, ✅ precios_proveedores.py, 🔶 competencia.py + m004; falta respaldo Gemini de precios) |
| 4 — Costeo | ⏳ |
| 5 — Financiero | ⏳ |

## 5. Carpeta en la PC

`C:\Users\franf\Desktop\ROBOT PANADERO` (comillas en comandos). PowerShell en VS Code.
Descargas: a `Escritorio\download` (nunca Descargas); de ahí con `Copy-Item`.
**Respaldo del Excel antes de la m003**: `Escritorio\respaldo_antes_m003`
(Excel descargado de Drive el 24-sep 17:45, 219 KB, nombre
`Estudio_Financiero_Panadero_CON_ROUND_v68_DRIVE_CALCULADO.xlsx`). No se toca nunca.
**Copia de prueba** = `datos_prueba\maestro_prueba.xlsx` (ya tiene la m003); la
usan todos los comandos con `--local`. Se refresca con `Copy-Item` desde un respaldo.
Archivos viejos de otras carpetas (bitácoras V2, etc.) se conservan; no se borran.

## 6. Entorno y comandos

- Python 3.14; usar `py -m pip ...` y `py -m <modulo>`. LibreOffice instalado.
- Probar conexión: `py -m conexion_drive.probar_conexion` (25 hojas)
- Pedidos (local): `py -m microservicios.grupo1_captura_operativa.probar_pedidos_digitales`
- **Precios de proveedores**: `py -m microservicios.grupo3_integraciones.precios_proveedores --solo-consultar`
  (solo consulta, no toca Excel) · `--local` (copia de prueba) · sin opción = Drive real.
- Migración m003: `py -m migraciones.m003_proveedores_3c --local` (sin `--local` = Drive;
  si ya está aplicada responde "Ya estaba aplicada").
- **Competencia** (Musmanni + Musi): `py -m microservicios.grupo3_integraciones.competencia --solo-consultar --sin-vision`
  (solo capa 1) · `--probar-vision <URL imagen>` (una imagen con Gemini) · `--solo-consultar` (todo, sin Excel)
  · `--local` · sin opción = Drive real. Cache de lecturas: `datos_prueba/cache_competencia.json`.
- Migración m004: `py -m migraciones.m004_competencia --local` (sin `--local` = Drive).
- Verificar que un archivo copiado es el nuevo: comparar **Length** (bytes) con
  `Get-Item "<ruta>" | Format-Table Name, Length, LastWriteTime` (el navegador renombra a `(1)`, `(2)`...).
- Ver commits: `git log --oneline -3` · Estado: `git status` (debe decir "working tree clean").
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
  20d2329 (escudo _xlfn v2), **62d0de8 (m003 + reporte), 8cd2d98 (m003: corregir --local),
  3e8f656 (precios_proveedores.py)**.
- `git status` → `git add .` → `git commit -m "..."` → `git push`.
- `git grep AIza` debe salir vacío (ninguna clave en el repo).
- Nunca editar en la web de GitHub.

## 8. Documentos y claves (fuera del repo)

- **Libro de mapas**: vigente **V10** (m003, Reglas #16 cambio en todos los puntos,
  #17 reemplazo de materia prima, #18 precios por API de catálogo; bloques de
  proveedores con EAN/SKU/slug; pendientes #17-#19; pestaña "Cambios V10").
  **Falta que el usuario pegue la clave de Gemini en NODO INTEGRAL E52 (celda amarilla).**
  Próximo: **V11** (documentar precios_proveedores.py, competencia.py, m004 y el modelo Gemini).
- **Claves en el `.env`** (valores en el `.env` y en el libro de mapas, NO aquí):
  - `GOOGLE_MAPS_API_KEY`
  - `LIMITE_CAMBIO_PRECIO=0.5` — ACTIVO desde la 1a corrida real (24-sep): cambio de
    precio >50% no se escribe y queda marcado en el evento.
  - `GEMINI_API_KEY` — proyecto "Default Gemini Project" (nivel gratuito,
    `projects/253837609158`). Formato nuevo de Google: empieza con `AQ.` (válido).
    La clave del proyecto con prepago (552251224672) daba 401: en desuso.
  - `GEMINI_MODELO` (OPCIONAL, no creado): fija el modelo. Si no está, competencia.py elige el
    "flash" más nuevo de la clave (25-sep: `gemini-3.8-flash`; `gemini-2.5-flash` ya da 404) y
    usa otro "flash" de respaldo si el elegido está saturado (503/429).
- Drive: "Conservar para siempre" no funcionó desde la cuenta del usuario
  (probablemente porque la versión la subió el robot) → respaldo manual a la PC.

## 9. Dónde quedamos (25-sep-2026, cierre ~01:00)

**Decisiones de hoy (competencia):**
- **Musi TAMBIÉN es competidor** (el usuario estuvo en el local: la mitad es panadería). Dos
  competidores: Musmanni y Musi. Facebook descartado como fuente (términos de uso, login, API cerrada).
- Enfoque académico: aunque hoy no haya promos vigentes, lo que se demuestra es que la conexión
  funciona y que una promo nueva se detectaría. "Cero ofertas vigentes" es un resultado válido.
- La oferta de la competencia se queda en **Tabla Combos Dinamicos** (fila 10 Musmanni, fila 11 Musi),
  como estaba en el diseño original: en el SaaS se ve junto a los combos propios para ajustarlos.
  No son combos propios (NODO C19/C20 solo leen las filas 5 a 8).
- Sin promo vigente: se escribe "Sin promocion vigente (fecha)" y valor 0. Si falla la conexión: no se toca.

**Hallazgos de las fuentes (APIs de WordPress abiertas, sin IA):**
- Musmanni `/promociones/`: TODO oculto (galerías 2023/2024, banners y botón con `elementor-hidden-*`).
  Pero las subpáginas `promociones-reposteria`, `-almuerzos` y `-pasteles` SÍ tienen 15 imágenes
  visibles (algunas subidas en jul-2025). API `promos`: 1 sola, de mar-2024.
- Musi `promociones-para-panaderia-y-reposteria`: 9 imágenes visibles (campaña MAYO-JUNIO-25).
- En el Excel, NODO C11 NO tenía la URL (texto provisional "valor /scraping web"); la URL vive en el
  código, el libro de mapas y DiccionarioDatos L252. Ninguna fórmula lee C11, D10/E10 ni D11/E11.

**competencia.py** (microservicios/grupo3_integraciones/, 28960 bytes, NO subido a GitHub todavía):
- Capa 1: API WordPress; solo imágenes VISIBLES; descarta logos y SVG.
- Capa 2: Gemini lee producto, precio y vigencia (JSON). Cache por imagen (se guarda en seguida).
  Avance en pantalla ("imagen n de N", reintentos) y modelo de respaldo si hay 503/429.
- Capa 3: VIGENTE / SIN_PROMO / FALLO. Sin fecha impresa: vigente solo si se subió hace <= 45 días
  (`DIAS_SIN_VIGENCIA`). E = precio más bajo vigente.
- Freno: si la m004 no está aplicada, se detiene ANTES de escribir cualquier celda.
- Pruebas OK en la PC: `--sin-vision` y `--probar-vision` (Musi DC-MAYO-JUNIO-25-01: "2X Repostería
  Dulce o Salada o Panes Especiales", ₡1,350, "Válido por tiempo limitado" → VENCIDA).
- `--solo-consultar` completo: INTERRUMPIDO (Ctrl+C) por Google saturado (503) y luego 429
  (límite del nivel gratuito). Se corrió con la versión anterior (sin cache inmediato ni respaldo).
  **Pendiente confirmar que la versión de 28960 bytes quedó copiada en el proyecto.**

**m004_competencia.py** (migraciones/, NO aplicada aún; el usuario la tiene en Escritorio\download):
- A) NODO B11 = "DATA OFERTA MUSMANNI" + fila 21 para Musi (A21:E21). Se usó la fila 21 porque
  insertar una fila correría C12:C20 (fórmulas y nodo_integral.py).
- B) Tabla Combos Dinamicos C11:E11 para Musi. C) notas "|| m004" en DiccionarioDatos L252, N252, N89.
- D) DiccionarioDatos filas 303 (NODO.PRECIOS_COMPETENCIA_MUSI) y 304 (MKT.COMPETENCIA.MUSI).
- Probada por Claude sobre el Excel de Drive (versión descargada 24-sep 23:48): 40 celdas, 2a corrida
  "Ya estaba aplicada"; comparación independiente de todo el libro: 0 cambios no anunciados,
  4,538 fórmulas, 0 errores. (Flujo caja diario fila 212 cambia sola por HOY(), no por la m004.)

**Siguientes pasos (en este orden):**
1. Confirmar competencia.py = 28960 bytes en el proyecto. Correr `--solo-consultar` cuando Google
   no esté saturado (esperar 1-2 h o al día siguiente por el 429).
2. m004: copiar a `migraciones\`, `--local`, luego Drive. Luego `competencia --local` y Drive.
3. git: add/commit/push de competencia.py, m004 y esta bitácora V7.
4. Libro de mapas V11 (precios_proveedores.py, competencia.py, m004, modelo Gemini). Pegar la clave
   Gemini en NODO INTEGRAL E52 del mapa (celda amarilla).
5. Respaldo de precios con ScrapeGraphAI + Gemini. 6. Grupos 4 y 5. 7. Bot de WhatsApp real.
8. Limpieza de copias viejas en Escritorio\download (con todo ya en GitHub).

**Pendientes para el Grupo 5:**
- Anular una línea de pedido digital: el registro original se mantiene y se
  agrega una **fila negativa** (asiento de reversión). Definir estado, fecha y envío.
- "Ventas digitales" de Flujo caja suma TOTAL_A_COBRAR, que incluye el envío.

**Pendientes de hosting:** minutos de Actions (ping a Render); respaldo diario
del Excel ("Keep forever" vía Drive API, límites de versiones). En la nube, el cache de
competencia.py se perdería entre corridas (definir dónde guardarlo).

## 10. Para continuar en un chat nuevo

Subir: **este archivo (V7)** y el **libro de mapas V10**. Si se va a aplicar la m004, no hace falta
subir el Excel (el script revisa cada celda antes de escribir). Decir:
"Continuemos con competencia.py y la m004 del Grupo 3C".
