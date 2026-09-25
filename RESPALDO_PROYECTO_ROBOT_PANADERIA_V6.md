# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V6

*Última actualización: 24-sep-2026 (cierre de sesión, ~21:00). Reemplaza a la V5. Guarda este archivo en la
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
se escriben en 2 TABLAS del Excel. Nunca decir "tiendas".
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
| 3 — Integraciones | ✅ 3A · ✅ 3B · 🔶 3C (✅ m003, ✅ precios_proveedores.py; falta Musmanni y respaldo Gemini) |
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
  Próximo: **V11** (documentar precios_proveedores.py y Musmanni).
- **Claves en el `.env`** (valores en el `.env` y en el libro de mapas, NO aquí):
  - `GOOGLE_MAPS_API_KEY`
  - `LIMITE_CAMBIO_PRECIO=0.5` — ACTIVO desde la 1a corrida real (24-sep): cambio de
    precio >50% no se escribe y queda marcado en el evento.
  - `GEMINI_API_KEY` — proyecto "Default Gemini Project" (nivel gratuito,
    `projects/253837609158`). Formato nuevo de Google: empieza con `AQ.` (válido).
    La clave del proyecto con prepago (552251224672) daba 401: en desuso.
- Drive: "Conservar para siempre" no funcionó desde la cuenta del usuario
  (probablemente porque la versión la subió el robot) → respaldo manual a la PC.

## 9. Dónde quedamos (24-sep-2026, cierre ~21:00)

**Hecho hoy:**
- **Migración m003 APLICADA al Excel real de Drive** (versión 60, mismo ID): leche Suli →
  "Leche Semidescremada Dos Pinos -1 L" (11 celdas), 7 materias primas de Solera
  renombradas en MP/Ingreso/Control (21), 6 conversiones col E, 15 precios verificados
  col D (opción D1), 5 notas en DiccionarioDatos. Doble revisión: 58 celdas exactas,
  0 cambios no anunciados, 4,538 fórmulas y 0 errores. Reporte: `migraciones/REPORTE_m003.md`.
  Incidente: la 1a versión del script no manejaba `--local` y se aplicó directo a Drive
  (sin daño: el pre-chequeo confirmó que Drive = respaldo). Corregido y probado.
- Efectos a revisar (pendientes #18 y #19 del mapa): combos 3 y 4 suben (queso/natilla
  tenían costo de relleno ₡100); Flujo caja diario L12 baja por compras de prueba con
  precios reales; costo baguette 483.30 → 411.23, cupcake 217.19 → 229.45.
- **precios_proveedores.py COMPLETO** (microservicios/grupo3_integraciones/): consulta las
  APIs de catálogo de 2 proveedores (Solera por slug, respaldo SKU; MaxiPalí por EAN,
  leche con respaldo Sabemas), 16 productos, escribe solo col D de TCM MP y TCM; si falla
  no toca la celda; evento PRECIOS_PROVEEDORES en RegistroEventos. Probado: 16 de 16,
  `--local` y Drive real OK ("Sin cambio: 16"). Límite 50% activo.
- Libro de mapas V10 entregado. VS Code limpio (se cerró sin guardar un .json de
  credenciales de la carpeta FUNIVER que aparecía vacío en el editor; el disco quedó intacto).

**Musmanni — hallazgos (24-sep, noche):**
- Musmanni sigue siendo la PANADERÍA (competidor correcto). **Musi** = tiendas de
  conveniencia de FIFCO (antes "Musmanni Mini Súper", desde 2017): otro rubro, no se usa.
- La página `https://musmanni.com/promociones-2/` (NODO C11) está **desactualizada**:
  modificada 22-mar-2024, una sola promo (Melcochón: "porción de pastel + café pequeño",
  martes y jueves). El precio NO está como texto (seguramente en la imagen).
- El menú real apunta a `https://musmanni.com/promociones/` con subpáginas
  `promociones-reposteria/`, `promociones-almuerzos/`, `promociones-pasteles/`.
- **API de WordPress ABIERTA**: `https://musmanni.com/wp-json/wp/v2/promos` devuelve JSON
  (id, fecha, título, slug, link, categoría, `featured_media` = id de la imagen). Hoy
  devuelve SOLO 1 promo (id 2663, 14-mar-2024). Imagen: `https://musmanni.com/wp-json/wp/v2/media/<featured_media>`.
  → Las promos vigentes podrían no estar en ese tipo de contenido.

**Siguientes pasos (en este orden):**
1. **Musmanni** (chat nuevo): el usuario manda captura de `https://musmanni.com/promociones/`
   (y sus 3 subpáginas si tienen promos) y, si hace falta, su código fuente.
   Decidir la fuente: API `promos` (sin IA para texto) + **Gemini solo para el precio
   dentro de la imagen**; o lectura de la página vigente. Anunciar plan antes de construir.
   Actualizar NODO C11 con la URL correcta (con migración si toca el Excel).
2. Respaldo de precios con ScrapeGraphAI + Gemini (cuando falle una API de catálogo).
3. Libro de mapas V11 (precios_proveedores.py, Musmanni).
4. Grupos 4 y 5. 5. Bot de WhatsApp real (bilingüe, función compartida).

**Pendientes para el Grupo 5:**
- Anular una línea de pedido digital: el registro original se mantiene y se
  agrega una **fila negativa** (asiento de reversión). Definir estado, fecha y envío.
- "Ventas digitales" de Flujo caja suma TOTAL_A_COBRAR, que incluye el envío.

**Pendientes de hosting:** minutos de Actions (ping a Render); respaldo diario
del Excel ("Keep forever" vía Drive API, límites de versiones).

## 10. Para continuar en un chat nuevo

Subir: **este archivo (V6)**, el **libro de mapas V10** y, para Musmanni, la
**captura de `musmanni.com/promociones/`** (+ `promos.json` si se quiere). El Excel
solo si se va a migrar algo (descargarlo de Drive sin abrirlo). Decir:
"Continuemos con Musmanni del Grupo 3C".
