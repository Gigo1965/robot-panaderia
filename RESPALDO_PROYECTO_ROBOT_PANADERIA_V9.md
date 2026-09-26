# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V9

*Última actualización: 25-sep-2026 (noche, ~10 p. m. CR). Reemplaza a la V8. La V8 completa sigue
abajo SIN CAMBIOS (sección "CONTENIDO V8"); lo nuevo está en este bloque "ACTUALIZACIÓN V9".
Subir este archivo y el libro de mapas V12 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

---

## ACTUALIZACIÓN V9 (25-sep-2026) — Grupo 4 cerrado, Grupo 5 muy avanzado

**Sesión de ~7 h (3:00 p. m. a 10:00 p. m. CR).** La lectura de Musi se pospuso al 26-sep
(cupo de Gemini agotado el 25-sep) para aprovechar el día con los Grupos 4 y 5.

### Commits nuevos (agregar a la lista de la sección 7)
- **6b315f3** — bitácora V8 (25-sep; la V8 no lo trae porque se escribió antes).
- **9fd8ddd** — `conexion_drive/probar_cache_competencia.py` (prueba solo lectura del cache en Drive: TODO BIEN).
- **629b035** — `migraciones/m005_costeo.py` + `migraciones/comparar_libros.py`.
- **87c2ae8** — `microservicios/grupo4_costeo/` (`__init__.py`, `costeo.py`, `probar_costeo.py`).
- **99d89a1** — `migraciones/m007_flujo_proyecto.py`.
- Pendiente: commit de esta bitácora V9 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V9.md`).

### Tareas manuales del usuario
- **Hecho:** `CACHE_COMPETENCIA.xlsx` subido a Drive y compartido con la cuenta de servicio;
  `probar_cache_competencia.py` = TODO BIEN (lectura y permiso de edición). Sigue POR CONFIRMAR la
  regla al llegar al límite (Decisiones Pendientes #20).
- Sin confirmar en este chat: clave de Gemini en el libro de mapas, NODO INTEGRAL E52 (#21).

### Herramienta nueva: `migraciones/comparar_libros.py` (4811 bytes)
Segunda revisión INDEPENDIENTE del protocolo (no usa el código de la migración): compara todas las
hojas celda por celda (fórmulas, valores, rangos de tabla, errores nuevos). Uso:
`py -m migraciones.comparar_libros "datos_prueba/maestro_prueba_ANTES_mXXX.xlsx" "datos_prueba/maestro_prueba.xlsx"`
Flujo de toda migración: foto previa `maestro_prueba_ANTES_mXXX.xlsx` (sin pisar si existe) →
`--local` → comparar_libros → leer valores clave → Drive → `probar_conexion` → git (archivo por nombre).

### m005_costeo (16231 bytes) — APLICADA (local y Drive v73)
Revisión celda por celda de la cadena API de catálogo → TCM MP/TCM → hojas de costo. Corregido:
- cupcakes C13 = H18 (margarina 10 g) y C17 = H16 (levadura 1.5 g): el usuario REEMPLAZÓ su
  decisión anterior ("diseño intencional"): "hay que corregir para que tenga lógica dentro de costos".
- cupcakes C31 MONTH($C$7) (antes leía la hora = siempre enero); H21 = TCM D10 (leche a costo, #3);
  G5 = 0.025 kg; G17 "Margarina 1gr".
- baguette C16 = H22 (bolsa kraft real; antes 5 fijo; 1 bolsa por baguette).
- valor activos: Tabla12 A2:H33 (20 filas de reserva con fórmulas IF), totales a filas 34-36;
  flujo caja proyecto D23 → D36.
- Local: A 7, B 90, C 1, D 12 celdas; comparar_libros: 4,527 fórmulas iguales, 0 errores
  (121 = 110 + 11 del evento EVT-000009). Costo unitario: baguette **430.23**, CupCake **248.05**;
  precios de referencia sin cambio (842.69 / 343.16).

### costeo.py (20594 bytes) — Grupo 4, probado 20/20 en la PC
Funciones: `actualizar_precio_referencia` (G28/G27; **bloquea** si queda bajo el costo unitario),
`actualizar_plaza` (planilla B/C filas fijas 7-12; "Otros 1"/"Otros 2"), `agregar_plaza` (filas 13-32),
`actualizar_gasto` (9 rubros manuales; patente/permiso repetido = aviso), `agregar_activo` /
`actualizar_activo` (baja = cantidad 0). Cada cambio: evento con valor anterior → nuevo.
Uso manual: `py -m microservicios.grupo4_costeo.costeo --local precio P01 850` (ver docstring).
Prueba: `py -m microservicios.grupo4_costeo.probar_costeo` (copia temporal `datos_prueba\_prueba_costeo.xlsx`).
La carpeta `grupo4_costeo` ya existía vacía. **GRUPO 4 CERRADO.**

### Grupo 5 — decisiones de método del usuario (para la tesis; Regla #25 del mapa)
- Punto de equilibrio = cubrir TODAS las obligaciones (operacional + deuda) + **20% de ganancia**;
  crecimiento anual = **inflación 3%**.
- **Costo de producción** (MP + mano de obra directa panadero/ayudante + CIF) ≠ **gasto administrativo**
  (administrador, contador, misceláneos); ambos suman el costo total de la empresa. `planilla` = total
  del staff (no se toca); cargas sociales completas en gasto operacional.
- Sueldo por hornada y costos de producción = **colchón** para imprevistos (semi industrial: la
  empresa genera más de lo reportado; economía de escala). Precio de referencia fijo.
- **Sin depreciación** en el flujo; **sin valor de rescate** (precio incierto del equipo usado y cambio
  tecnológico). Costo del proyecto = hornada estimada (#5 cerrado). Fila 16 del flujo vacía a propósito.
- Regla del usuario: **no quitar nada del modelo sin preguntar** (Regla #23). Consejo dado: presentarlo
  como "modelo conservador"; ante Hacienda se declara la utilidad real.

### m007_flujo_proyecto (11193 bytes) — APLICADA (local y Drive v76)
- A) gasto operacional B9:M9 "Planilla" = solo salarios administrativos (M14/M15 − 95% del sueldo de
  panadero y ayudante ya en el costo). Enero 371,500 (antes 970,000). Resuelve #6.
- B+C) flujo caja proyecto 2026-2030 desde sus propias filas, sin ×12, 3% inflación, utilidad neta =
  antes de impuestos − impuesto (antes no lo restaba).
- E) Punto de Equilibrio B23 = (B15+B21)×B22 (20% sobre gasto + deuda).
- Local: A 12, B 19, C 5, D 8, E 1; comparar_libros: 4,580 fórmulas iguales, 56 distintas, 0 errores
  (evento EVT-000010). **Resultado: VAN ₡2,446,342; TIR 20.9%; recuperación 3.03 años**
  (tasa del préstamo 12.51%); 46,054 unidades por producto/año.

### Pendientes (en este orden, próximo chat)
1. **Lectura real de Musi** (cupo reiniciado): `py -m microservicios.grupo3_integraciones.competencia --local`,
   luego sin `--local` y `py -m conexion_drive.probar_conexion` → **cierra el Grupo 3**.
2. Grupo 5, pequeños (~2.5-3.5 h): **#18** datos de prueba de Ingreso de Inventario; **#15** envío dentro
   de "Ventas digitales" (decidir si va aparte); **#7** anular línea de pedido digital = fila negativa
   (adjuntar `microservicios/grupo1_captura_operativa/pedidos_digitales.py`, 19342 bytes).
3. Grupo 5, decisiones del usuario/tutor: **F6** tasa de impuesto (7% renta vs "IVA 1%" en flujo anual, #23);
   **F7** financiamiento (B32, C30, D26, E21:I22, TIR VAR C10; #24).
4. **m006**: café a la venta por paquete de 250 g (mismo nombre técnico que Control de Inventario fila 21;
   compras ×10; #25).
5. Notas de metodología en la tesis (#26). Revisar `romper_prueba.py` (VS Code intentó ejecutarlo con el
   botón Run; no se ejecutó; #27).
6. Después: bot de WhatsApp real (Twilio + plantillas Meta, iniciar trámite temprano) y hosting.

### Tips aprendidos hoy
- Indicar SIEMPRE carpeta (ROBOT PANADERO o download) y dónde se escribe (Terminal o ventana "Abrir").
- Adjuntar archivos: `Set-Clipboard "C:\Users\franf\Desktop\download"` → botón adjuntar del chat →
  ventana "Abrir": Ctrl+V en "Nombre de archivo" + Enter → escribir `"a.py" "b.py"` + Enter.
- Copiar sin pisar: `if (Test-Path "destino") { "YA EXISTE" } else { Copy-Item ...; "COPIADO" }`.
- `New-Item` da error rojo inofensivo si la carpeta ya existe; revisar contenido con `Get-ChildItem -Force`.
- Si en la Terminal aparece `& '...archivo.py'` sin haberlo escrito, es el botón Run de VS Code: Esc.

---

# CONTENIDO V8 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V8

*Última actualización: 25-sep-2026 (tarde). Reemplaza a la V7. Guarda este archivo en la
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
Nunca se borran archivos de la PC sin pedirlo y explicarlo (antes de borrar: `-WhatIf`).
**Vocabulario (precisión para la tesis)**: Solera y MaxiPalí son PROVEEDORES; cada
uno tiene una API DE CATÁLOGO (fuente de datos JSON); se consultan 16 PRODUCTOS y
se escriben en 2 TABLAS del Excel. Nunca decir "tiendas". **Musi es el COMPETIDOR
monitoreado** (Musmanni desactivado; ver sección 9). Ambas marcas son de **Heineken
Costa Rica desde ene-2026** (antes FIFCO).
**Tip de Terminal**: si al pegar un comando aparece un `+` al inicio (`+py ...`), da error.
Antes de pegar, clic en la terminal y **Esc** (borra la línea).
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
   La competencia y los precios de proveedores corren en la MISMA tarea diaria
   (7:00 a. m. hora CR = 13:00 UTC, en la fase de hosting).
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
| 3 — Integraciones | ✅ 3A · ✅ 3B · 🔶 3C (✅ precios_proveedores.py, ✅ competencia.py + m004 en Drive; **falta solo la 1a lectura real de Musi** con cupo nuevo de Gemini) |
| 4 — Costeo | ⏳ siguiente |
| 5 — Financiero | ⏳ (de una vez después del 4) |

Respaldo de precios con ScrapeGraphAI + Gemini: **DESCARTADO** por el usuario (25-sep-2026),
por razones técnicas (las APIs de catálogo dan el precio exacto sin IA, 16 de 16) y económicas
(el cupo gratuito de Gemini no alcanza para dos usos).

## 5. Carpeta en la PC

`C:\Users\franf\Desktop\ROBOT PANADERO` (comillas en comandos). PowerShell en VS Code.
Descargas: a `Escritorio\download` (nunca Descargas); de ahí con `Copy-Item`.
El navegador renombra con `(1)`, `(2)`...: **identificar siempre por tamaño (Length)** antes de copiar.
**`download` quedó limpia el 25-sep** (todo ya estaba en GitHub).
**Respaldo del Excel antes de la m003**: `Escritorio\respaldo_antes_m003` (no se toca nunca).
**Copia de prueba** = `datos_prueba\maestro_prueba.xlsx` (ya tiene m003 y **m004**); la usan todos los
comandos con `--local`. Foto previa a la m004: `datos_prueba\maestro_prueba_ANTES_m004.xlsx`.
Archivos viejos de otras carpetas (bitácoras V2, etc.) se conservan; no se borran.

## 6. Entorno y comandos

- Python 3.14; usar `py -m pip ...` y `py -m <modulo>`. LibreOffice instalado.
- Probar conexión: `py -m conexion_drive.probar_conexion` (25 hojas; 25-sep: versión 70 en Drive)
- Pedidos (local): `py -m microservicios.grupo1_captura_operativa.probar_pedidos_digitales`
- **Precios de proveedores**: `py -m microservicios.grupo3_integraciones.precios_proveedores --solo-consultar`
  (solo consulta, no toca Excel) · `--local` (copia de prueba) · sin opción = Drive real.
- **Competencia (solo Musi)**: `py -m microservicios.grupo3_integraciones.competencia --solo-consultar --sin-vision`
  (solo capa 1) · `--probar-vision <URL imagen>` (una imagen con Gemini) · `--solo-consultar` (todo, sin Excel)
  · `--local` · sin opción = Drive real. Cache: `datos_prueba/cache_competencia.json`.
- Migraciones: `py -m migraciones.m003_proveedores_3c` / `py -m migraciones.m004_competencia`
  (`--local` = copia de prueba; ambas ya aplicadas en Drive: responden "Ya estaba aplicada").
- Verificar un archivo copiado: `Get-Item "<ruta>" | Format-Table Name, Length, LastWriteTime`.
- Listar download: `Get-ChildItem "$env:USERPROFILE\Desktop\download\*<nombre>*" | Format-Table Name, Length, LastWriteTime`
- Borrar con ensayo previo: `Remove-Item "<ruta1>", "<ruta2>" -WhatIf` (muestra, no borra) y luego sin `-WhatIf`.
- Ver commits: `git log --oneline -3` · Estado: `git status` (debe decir "working tree clean").
- Google Maps: `py -m microservicios.grupo3_integraciones.google_maps <lat> <lng>`
- Romper copia de prueba (test escudo): `py romper_prueba.py`
- Recalcular copia local a mano:
  `py -c "from pathlib import Path; from conexion_drive.recalculo import reparar_prefijos_xlfn, recalcular_bytes; p=Path('datos_prueba/maestro_prueba.xlsx'); c,r=reparar_prefijos_xlfn(p.read_bytes()); p.write_bytes(recalcular_bytes(c)); print('OK')"`
- Probar clave Gemini (debe dar 200):
  `py -c "import os,http.client as h; from conexion_drive import config; c=h.HTTPSConnection('generativelanguage.googleapis.com',timeout=20); c.request('GET','/v1beta/models',headers={'x-goog-api-key': os.getenv('GEMINI_API_KEY')}); r=c.getresponse(); print(r.status, r.read()[:200].decode())"`

## 7. Git / GitHub

- Repo privado https://github.com/Gigo1965/robot-panaderia.git (main).
- Commits: V4, cd7db42 (G2), c52e289 (NODO v3), 67d3361 (escudo), 5becbcc,
  a3a69f1 (3B parte 1), e0c529d (3B parte 2), 20d2329 (escudo _xlfn v2),
  62d0de8 (m003 + reporte), 8cd2d98 (m003 --local), 3e8f656 (precios_proveedores.py),
  0625efb (bitácora V6), **29fb71d (competencia.py versión inicial + bitácora V7)**,
  **41521ba (competencia.py solo Musi + cupo diario + lectura incompleta; m004)**.
- `git status` → `git add .` → `git commit -m "..."` → `git push`.
- El aviso "LF will be replaced by CRLF" es inofensivo (fin de línea de Windows).
- `git grep AIza` debe salir vacío (ninguna clave en el repo).
- Nunca editar en la web de GitHub.

## 8. Documentos y claves (fuera del repo)

- **Libro de mapas**: vigente **V11** (25-sep-2026): bloques V11 en NODO INTEGRAL, Tabla combos
  dinamicos, Mapa TCM MP y Mapa TCM (precios_proveedores.py); Reglas #19 (lectura incompleta =
  no se toca) y #20 (solo Musi); Decisiones Pendientes #20-#22; pestaña "Cambios V11".
  Revisión independiente V10→V11: 44 hojas (1 nueva), 6,200 celdas iguales, 15 anotadas con
  "|| V11", 180 nuevas, 0 borradas, comentarios intactos.
  **Falta que el usuario pegue la clave de Gemini en NODO INTEGRAL E52 (celda amarilla).**
- **Claves en el `.env`** (valores en el `.env` y en el libro de mapas, NO aquí):
  - `GOOGLE_MAPS_API_KEY`
  - `LIMITE_CAMBIO_PRECIO=0.5` — ACTIVO desde la 1a corrida real (24-sep).
  - `GEMINI_API_KEY` — proyecto "Default Gemini Project" (nivel gratuito, `projects/253837609158`).
    Formato nuevo de Google: empieza con `AQ.` (válido).
  - `GEMINI_MODELO` (OPCIONAL, no creado). Si no está, competencia.py elige el "flash" más nuevo
    (25-sep: `gemini-3.8-flash`; respaldo 3.7 y 3.6, en rueda). `gemini-2.5-flash` ya da 404.
- Drive: "Conservar para siempre" no funcionó desde la cuenta del usuario → respaldo manual a la PC.

## 9. Dónde quedamos (25-sep-2026, tarde)

**Decisiones de hoy:**
- **Solo se monitorea a Musi.** Musi nació en 2017 como el nuevo nombre de los Musmanni Mini Súper;
  las panaderías Musmanni siguen como marca aparte (204 franquicias; sin promos web nuevas desde
  jul-2025). Ambas marcas son de **Heineken Costa Rica** (compra a FIFCO cerrada en ene-2026).
  Musmanni queda **desactivado** en competencia.py (`"activo": False`, reversible); sus celdas
  (NODO C11, Tabla Combos D10/E10) no se tocan.
- **Horario**: 1 vez al día, **inicio 7:00 a. m. hora CR** (13:00 UTC) para tener el resultado
  seguro a las 8:00, en la misma tarea programada del NODO INTEGRAL (fase de hosting).
- **Regla #19 — lectura incompleta = no se toca**: si Gemini no pudo leer alguna imagen visible,
  el competidor queda FALLO y sus celdas conservan el último dato (con su fecha).
- ScrapeGraphAI descartado (sección 4).

**competencia.py (32619 bytes, commit 41521ba):**
- Capa 1: API WordPress de Musi (9 imágenes visibles, 0 ocultas). Capa 2: Gemini con cache
  inmediato por imagen. Capa 3: VIGENTE / SIN_PROMO / FALLO.
- Modelo "flash" más nuevo en **rueda**; detecta el **cupo DIARIO** (429 con "PerDay"): marca el
  modelo sin cupo y pasa a otro; si ninguno tiene, deja de llamar a Gemini en esa corrida.
- Probado hoy en la PC: `--solo-consultar` y `--local` → cupo diario agotado en los 3 modelos →
  Musi FALLO (lectura incompleta 0 de 9) → **celdas intactas**, evento registrado. Correcto.
- Hallazgo para la tesis: el cupo gratuito diario de los modelos nuevos es muy bajo; el cache es
  pieza obligatoria del diseño (en operación normal Gemini solo lee imágenes NUEVAS; corrida ~1 min).
- Lecturas ya en cache (Musmanni, jul-2025, todas VENCIDAS o no promo): 3 pancitos dulces ₡1,000;
  4 rollos de canela ₡2,500; sándwich + bebida ₡1,400.

**m004_competencia.py (12601 bytes) — APLICADA:**
- Local: 40 celdas (A 6, B 3, C 3, D 28) + evento; 2a corrida "Ya estaba aplicada".
  Comparación independiente de todo el libro: 25 hojas, 4,538 fórmulas iguales, 0 errores,
  51 cambios = 40 + 11 del evento; C19/C20 idénticas; formato fila 21 = fila 11.
  (Flujo caja diario fila 212 cambia sola por HOY().)
- Drive: mismos 40 cambios, "Ya estaba aplicada" en la 2a corrida, probar_conexion OK (versión 70).
- Notas corregidas antes de aplicar: modelo Gemini (ya no "2.5 Flash"), horario 7:00, lectura
  incompleta, Musi/Musmanni y Heineken (DiccionarioDatos N252, N89, fila 303).

**Siguientes pasos (en este orden):**
1. **Mañana en la mañana** (cupo de Gemini reiniciado ~1:00 a. m. CR):
   `py -m microservicios.grupo3_integraciones.competencia --local` → debe decir SIN_PROMO o VIGENTE
   para Musi; luego sin `--local` (Drive). Con eso **el Grupo 3 queda cerrado**. Si vuelve a salir
   lectura incompleta, no hay que cambiar código: se repite más tarde.
2. Usuario: pegar la clave de Gemini en el libro de mapas V11, NODO INTEGRAL E52.
   Usuario: subir `CACHE_COMPETENCIA.xlsx` a la carpeta del Excel maestro en Drive (sin convertir),
   compartirlo con la cuenta de servicio como Editor y anotar su ID (lo usa la fase de hosting).
3. git: add/commit/push de esta bitácora V8.
4. **Grupo 4 — Costeo** (costo pan baguette, costo cupcakes, planilla, gasto operacional, valor
   activos). No depende de Gemini.
5. **Grupo 5 — Financiero**, de una vez después del 4.
6. Bot de WhatsApp real (bilingüe) y hosting (Render + GitHub Actions). Iniciar temprano los
   trámites de Twilio/Meta (aprobación de plantillas puede tardar días).

**Estimación (25-sep-2026):** 8 a 12 sesiones, 30 a 45 h, unas 2 a 3 semanas al ritmo actual.
Grupo 4: 1-2 sesiones; Grupo 5: 3-4 (el más incierto); bot real: 2-3; hosting: 1-2; cierre y
README para tutores: 1. Power BI fuera de la cuenta (después de datos reales).

**Pendientes para el Grupo 5:**
- Anular una línea de pedido digital: el registro original se mantiene y se
  agrega una **fila negativa** (asiento de reversión). Definir estado, fecha y envío.
- "Ventas digitales" de Flujo caja suma TOTAL_A_COBRAR, que incluye el envío.

**Pendientes de hosting:**
- Minutos de Actions (ping a Render); respaldo diario del Excel ("Keep forever" vía Drive API).
- Programar la tarea de las 7:00 (NODO + competencia + precios); clima/tipo de cambio cada 30 min.
- **Cache de competencia en la nube** (Decisiones Pendientes #20). **DECIDIDO por el usuario:**
  un Excel real en Google Drive, **`CACHE_COMPETENCIA.xlsx`** (plantilla entregada el 25-sep):
  hoja `Cache` con la tabla real **TablaCache** (14 columnas: ID_LECTURA, FECHA_LECTURA, COMPETIDOR,
  URL_IMAGEN = llave, FECHA_SUBIDA_IMAGEN, ES_PROMOCION, PRODUCTO, PRECIO_CRC, VIGENCIA_DESDE,
  VIGENCIA_HASTA, TEXTO_VIGENCIA, ESTADO_AL_LEER, MODELO_GEMINI, ORIGEN); historial ACUMULADO
  (una fila por lectura; si una imagen se relee vale la más reciente). Hoja `Control`: límite del
  ciclo 1,000,000 filas (B4, editable), filas usadas, % usado, imágenes distintas, ciclo actual.
  Hoja `Instrucciones`. Subido por el usuario a la carpeta del Excel maestro, compartido con la
  cuenta de servicio como Editor; ID en el `.env` como `CACHE_COMPETENCIA_FILE_ID`.
  **Nunca abrirlo con Google Sheets** (Drive > Configuración: NO convertir archivos subidos).
  **POR CONFIRMAR:** qué pasa al llegar al límite. Propuesta: copia del ciclo en Drive
  (`CACHE_COMPETENCIA_CICLO_1.xlsx`) y ciclo nuevo conservando la última lectura de cada imagen
  (un reinicio total obligaría a Gemini a releer todo). En la práctica el límite no se alcanza.
  Falta: código en competencia.py para leer/escribir este Excel y migrar las lecturas del JSON
  (ORIGEN = "migrado JSON").

## 10. Para continuar en un chat nuevo

Subir: **este archivo (V8)** y el **libro de mapas V11**. Decir:
"Continuemos: lectura real de Musi (competencia --local y Drive) y luego Grupo 4 (Costeo)".
Para el Grupo 4 conviene subir también el Excel maestro más reciente (descargado de Drive).
