# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V13

*Última actualización: 28-sep-2026 (noche, ~10:50 p. m. CR). Reemplaza a la V12. La V12 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V12", que a su vez trae la V11, la V10, la V9 y la V8); lo nuevo está en este bloque "ACTUALIZACIÓN V13".
Subir este archivo y el libro de mapas V15 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: el trabajo fue el 28-sep-2026 (~6:15 a 10:50 p. m. en Costa Rica). Algunos textos del código y de
DiccionarioDatos dicen "27-sep-2026" (código del #32, escrito el 26-sep noche, fecha UTC) o "28-sep-2026". Horas de Drive
en UTC (v97: 29-sep 02:57 UTC = 28-sep 8:57 p. m. CR).*

---

## ACTUALIZACIÓN V13 (28-sep-2026, noche, ~4.5 h) — #32 resuelto (m012, Drive v97), café Ca01, F6/F7, #27, #29, Regla #34

### Estado de los 5 grupos
| Grupo | Estado |
|---|---|
| 1 — Captura operativa | ✅ COMPLETO (#32 compras con costo congelado; café Ca01 rechazado en compras) |
| 2 — Inventario | ✅ COMPLETO (Regla #34: no hay hornada sin materia prima) |
| 3 — Integraciones | ✅ CERRADO (quedan #14 y #20 para hosting) |
| 4 — Costeo | ✅ COMPLETO |
| 5 — Financiero | ✅ COMPLETO en el modelo: F6 (#23) y F7 (#24) resueltos sin cambios. Solo queda #26 (redacción de la tesis) |
**5 de 5 grupos completos en código.** Siguiente fase (otro chat): bot de WhatsApp real + hosting.

### Commits de este chat (todos con push)
- **e5b0aa3** — bitácora V12 (hecho al inicio del chat; antes aa770dd).
- **0b2702c** — #32 + café: `microservicios/grupo1_captura_operativa/ingreso_inventario.py` (**20588**),
  `probar_ingreso_inventario.py` (**7114**), `migraciones/m012_congelar_costo_compras.py` (**9123**, nuevo).
  3 files changed, 426 insertions, 22 deletions.
- **8c65135** — #27: `romper_prueba.py` (**2736**, raíz). 1 file, 60 insertions, 26 deletions.
- **9efa7f1** — Regla #34 + #29: `microservicios/grupo2_inventario/hornadas.py` (**17010**) y `probar_hornadas.py`
  (**7730**). 2 files, 130 insertions, 7 deletions.
- **Último commit de código: 9efa7f1.** Pendiente al cierre: commit de esta bitácora V13 (solo este archivo).
- **Drive: versión 97** (m012). probar_conexion TODO BIEN: 25 hojas, 8 productos, precios de venta sin cambio
  (Baguette 842.69, CupCake 343.16, Cf01 2,179.275), LibreOffice encontrado.

### Respaldos y fotos (nada se borró)
- Escritorio\download: `ingreso_inventario_ANTES_cafe.py` (16941) y `probar_ingreso_inventario_ANTES_cafe.py` (5934) = las
  versiones del #32 que pasaron la 1a prueba; `romper_prueba_ANTES_27.py` (1240); `probar_hornadas_ANTES_29.py` (4415,
  original); `probar_hornadas_ANTES_34.py` (4927, versión #29 que nunca tuvo commit); `hornadas_ANTES_34.py` (13289).
  Siguen los `_ANTES_32` del chat anterior (12543 / 3096). En download también hay versiones viejas de otros días
  (p. ej. `hornadas.py` 13019 del 23-sep): no estorban porque todo se copia POR TAMAÑO.
- datos_prueba: `maestro_prueba_ANTES_prueba32.xlsx` (248943), `maestro_prueba_DESPUES_prueba32.xlsx` (250262, evidencia de
  las pruebas del #32 y del café, filas 38-45), `maestro_prueba_ANTES_m012.xlsx` (247159 = Drive v94),
  `maestro_prueba_ANTES_prueba29.xlsx` (247427), `maestro_prueba_RESPALDO_20260928_214812.xlsx` (247427, creado por
  romper_prueba.py). **`maestro_prueba.xlsx` al cierre = 247427 (8:32:58 p. m.) = m012 aplicada = igual a Drive v97.**

### 1) #32 — Partes A y B probadas (ingreso_inventario.py)
- Prueba local TODO OK dos veces. Harina E = 17,510 y leche 1,050 como NÚMERO; F = D×E sigue fórmula; café Cf01 2,075.5 y
  L "Compras Externas Netas"; costos 0/None/−5/"abc" rechazados; hornada con costo None (producción sigue fórmula).
- Corrección de la prueba: la leche vigilada usaba un nombre viejo ("Leche Suli…"); ahora "Leche Semidescremada Dos
  Pinos -1 L" (TCM fila 10). Aparece y sube +12.

### 2) Café Ca01 / Cf01 — DECIDIDO (opción A) = Decisiones Pendientes #34
- Mismo nombre en la TCM (Cf01, paquete 2,075.5) y en la MP (Ca01, caja de 10, 20,755). "Ca01, 1" se contaba en silencio
  como 1 paquete.
- Decisión del usuario: **Ca01 se rechaza en compras** (`UsarCodigoDeVenta`, hereda de `ProductoNoReconocido`):
  "Registre la compra con Cf01 … (1 unidad de 'Ca01' = 10 unidades de Cf01). No se escribio nada." El 10 se calcula
  (20,755 / 2,075.5) solo para el mensaje; el robot NUNCA convierte cantidades. Autocompletado y sugerencias solo
  muestran Cf01. La MP fila 19 no se toca. Regla general: aplica a cualquier producto que esté en las dos tablas.
  Descartadas: B (convertir ×10 automático) y C (solo documentar). Prueba: paso 9 (Ca01 y ca01 rechazados; no se
  escribe fila). Queda en la Regla #32 || V15 (no es regla aparte, decisión del usuario).
- `git grep` antes de copiar: ninguna otra parte del código usa leer_productos_inventario / resolver / buscar_opciones.

### 3) #32 Parte C — m012_congelar_costo_compras — APLICADA (local y Drive v97)
- Revisa antes de escribir: TablaIngresos A3:L37; exactamente 18 compras y 16 sin compra; fórmula vieja exacta en E;
  valor > 0 e IGUAL al costo actual de la TCM por nombre. Bloque A: E4:E20 y E37 → su valor (1,050 … 17,510 harina …
  2,075.5 café … 840 natilla). Bloque B: DiccionarioDatos H49 y N49 (texto agregado; fórmula vieja sin el "=").
  Evento EVT-000017 (local). Marca de "ya aplicada": "V15 m012" en N49.
- Local y Drive: 18 filas, Bloque A 18, B 2; 2a corrida "Ya estaba aplicada" en los dos.
- comparar_libros: 5,035 fórmulas iguales, 31 celdas, **10 "valores calculados que cambiaron"**, 0 errores.
  **Los 10 NO son de la m012**: 'Flujo caja diario' agrega un día por fila hasta HOY() (A = A anterior + 1 mientras
  < TODAY()); la foto previa se calculó el 26-sep y la copia el 28-sep → filas 214 (27-sep) y 215 (28-sep), columnas
  A, D, H, I, J, sin movimientos; acumulado J igual 969,522.97. Verificado celda por celda con los dos archivos.
- Valores clave antes = después: Flujo 2-mar I5 −359,333.13, E5 470,045, F5 81,955; acumulado 969,522.97; Ingreso F11
  280,160 y F19 20,755; Control F21 10; TIR VAR C10 10,950,907.94; PE B29 46,054. **La m012 cambió 0 números.**
- Resultado: un cambio de precio del proveedor (7:00) ya no mueve las compras registradas (auditoría). **Regla #33.**

### 4) F6 (#23) y F7 (#24) — RESUELTOS SIN CAMBIOS
- **#23**: la decisión se había tomado el 26-sep noche (chat "AVANCE V8") pero NO quedó en la V12 ni en el mapa; se
  registró hoy. Los impuestos del modelo son solo REFERENCIA (7% renta en 'flujo caja proyecto' fila 19; 1% en 'flujo
  caja anual' fila 31). La tesis es académica: demuestra el ecosistema integral, no un análisis tributario. El precio
  de referencia ya incluye el IVA del pan (1%, canasta básica) y todo lo comprado en CR trae el IVA incluido. No se
  modifica ninguna fórmula.
- **#24**: el modelo queda como está (el préstamo B32 = D23 + D24 financia el 100%; año 0 resta la inversión y los años
  1-5 la cuota). El usuario considera que no hay doble conteo; **no queda abierto para el tutor**. Se explica en la
  tesis como supuesto conservador.

### 5) #27 — romper_prueba.py — RESUELTO (opción B), commit 8c65135
- Qué es: evidencia del escudo _xlfn (#11, commit 67d3361). Rompe a propósito NODO INTEGRAL C19/C20 (quita `_xlfn.` de
  las dos TEXTJOIN) SOLO en la copia local; nunca toca Drive.
- Dos seguros nuevos: (1) pide escribir SI (acepta mayúsculas o minúsculas) → un Run accidental no hace nada;
  (2) el respaldo nunca se pisa (si `maestro_prueba_RESPALDO.xlsx` existe, el nuevo lleva fecha y hora).
- Probado en la PC: con "si" rompió las 2 celdas y creó `_RESPALDO_20260928_214812` sin pisar el viejo (229529, 24-sep);
  se restauró la copia desde ese respaldo; con "no" dijo "Cancelado. No se toco nada." y la copia no cambió.

### 6) #29 — probar_hornadas.py — RESUELTO, commit 9efa7f1
- Etiquetas "cantidad Baguette / CupCake (mes)" (Regla #26). Además (hallazgo): levadura y leche vigiladas con nombres
  viejos ("LEVADURA TEMPO SALADA 500 GR", "Leche Suli…") → no se mostraban; ahora "LEVADURA FRESCA GOLONDRINA 400 GR CIP"
  (Control fila 17) y "Leche Semidescremada Dos Pinos -1 L" (fila 7). Números esperados de levadura con paquete de 400 g
  ('Consumo de Materia Prima' F9/F15): 474 g = **1.185** paquetes (antes 0.948) y **0.15** (antes 0.12).
- Las notas viejas de DiccionarioDatos quedan como historia (si se tocan, con migración y "|| Vxx").

### 7) Regla #34 — No hay hornada sin materia prima (#35, opción C), commit 9efa7f1
- Hallazgo: la prueba de hornadas dejó existencias NEGATIVAS (harina 0.77 sacos − 1.144 = −0.374; leche −0.05). El
  usuario: "no puede haber hornada sin harina" → **opción C: TODA la receta, incluidos empaques** (bolsas, burbuja).
  Descartadas: A (solo harina) y B (sin empaques).
- `hornadas.py`: paso "0" dentro de `registrar_hornada`, antes de escribir (después de revisar duplicados):
  `_revisar_existencias` compara TablaRecetaConsumo (cantidad × por unidad / contenido por paquete) con la Existencia
  Actual de 'Control de Inventario' (columna F, paquetes). Si algo no alcanza → `MateriaPrimaInsuficiente` (hereda de
  HornadaInvalida) con TODOS los faltantes (paquetes y g/ml/und). `anular_hornada` no cambia (nunca se bloquea).
  `git grep`: solo hornadas.py, probar_hornadas.py y un comentario en `__init__.py` usan registrar_hornada.
- `probar_hornadas.py`: paso 0 (120 Baguette sin materia prima → rechazo con 4 faltantes: harina, mejorante, levadura,
  bolsas; nada escrito, Historico max_row 998) y paso 0b (compra de prueba Ha01 2, Me01 1, Le01 2, Bo01 2, L01 1 con
  ingreso_inventario, costo congelado). Luego pasos 1-7: harina −1.144 / levadura −1.185 / leche −0.6; tras anular
  −0.04 / −0.15 / −0.6; sin existencias negativas. **RESULTADO: TODO OK (0 lineas ERROR).**
- **Efecto en Drive v97:** bolsas kraft = 0, harina 0.77 sacos, mejorante 0.1 kg → **no se puede registrar ninguna
  hornada de Baguette hasta registrar compras** (con ingreso_inventario, no con migración). Para la demostración del
  bot/SaaS habrá que registrar primero compras de prueba en Drive (pendiente).
- Existencias Drive v97 (Control F): harina 0.77, mejorante 0.1, levadura 0.6999, bolsas 0, margarina Tumba 0.0934,
  leche 0.55, azúcar 0.1063, sal 0.7241, manteca 0.7241, burbuja 0.28.

### 8) Libro de mapas V15 (LIBRO_MAPA_DE_CELDAS_PFV15.xlsx, **244589** bytes) — FUERA DEL REPO
- **NO va a GitHub**: guarda claves de APIs (Regla #15; NODO INTEGRAL y Tabla combos dinamicos). Vive en la PC
  (y se adjunta al chat). En el chat se dijo por error "se sube al cierre con la bitácora"; se corrigió al revisar.
- Contenido: Regla #33 nueva (costo de compra congelado); Regla #32 || V15 (café: código de venta; Ca01 rechazado);
  **Regla #34 nueva** (no hay hornada sin materia prima); Decisiones Pendientes: #23, #24, #27, #29, #32 RESUELTOS;
  bloque "ACTUALIZACIÓN V15" con #34 (café) y #35 (hornada sin materia prima), ambos resueltos; Mapa_Ingreso Inventario
  fila 8 (D8/F8); Diseño del robot (4) D11 y A16; Mapa Consumo MP A17; pestaña "Cambios V15" (15 filas).
- Verificado contra V14 celda por celda: 0 borradas, 0 cambiadas, solo texto agregado y celdas/pestaña nuevas;
  comentarios y tablas intactos. Versiones intermedias descartadas: 242423 y 243086.

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V13 si no quedó hecho (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V13.md`).
2. **#26** notas de metodología en la tesis (redacción del usuario): colchón de costos, sin valor de rescate, sin
   depreciación, impuestos como referencia (#23), préstamo al 100% (#24), 20% sobre el punto de equilibrio, #30.
3. Datos de prueba en Drive: registrar compras de prueba de materia prima (bolsas, harina, mejorante, levadura…) antes
   de demostrar hornadas (Regla #34). Decidir cantidades (Regla #27).
4. Menores: #4 (dos líneas idénticas en una factura), #21 (clave de Gemini en el mapa, acción del usuario), #17, #19, #22.
5. **Fase siguiente (otro chat):** bot de WhatsApp real y bilingüe (Twilio + plantillas de Meta; traductor del JSON #33;
   mensaje de confirmación #28; forma de pago en la comanda; recibo al entregar; nombres en inglés #16) y hosting
   (Render + GitHub Actions a las 7:00 con NODO + competencia + precios; cache de competencia en Drive #20; respaldo
   diario del Excel; endpoint de anulación solo para administración; programación NODO #14).

### Tips aprendidos hoy
- El navegador cambia nombres al descargar ("(1)", "(2)" o un espacio antes de ".py"/".xlsx"). Copiar SIEMPRE buscando
  por tamaño: `Get-ChildItem … | Where-Object Length -eq N` y `Copy-Item -LiteralPath`; el comando se detiene si no
  hay exactamente un archivo con ese tamaño.
- `git --no-pager grep` muestra las líneas del archivo que está hoy en la carpeta (sirve para confirmar qué versión hay).
- Una prueba que escribe en la copia local: foto antes → prueba → restaurar la copia (verificar tamaño y hora).
- comparar_libros cuenta como "valores calculados que cambiaron" las filas nuevas que 'Flujo caja diario' agrega por
  HOY(); no son efecto de la migración (revisar cuáles son antes de ir a Drive).
- Antes de subir un archivo a GitHub, revisar que no tenga claves (el libro de mapas las tiene: queda fuera del repo).

---

# CONTENIDO V12 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V12

*Última actualización: 28-sep-2026. Reemplaza a la V11. La V11 completa sigue abajo SIN CAMBIOS (sección
"CONTENIDO V11", que a su vez trae la V10, la V9 y la V8); lo nuevo está en este bloque "ACTUALIZACIÓN V12".
Subir este archivo y el libro de mapas V14 al inicio del próximo chat (el mapa V15 se hará después de la m012).*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: el trabajo fue el 26-sep-2026 por la noche en Costa Rica (~10:10 a 11:40 p. m.); los textos
escritos en el código dicen "27-sep-2026" (fecha UTC del servidor). Es el mismo día.*

---

## ACTUALIZACIÓN V12 (26-sep-2026, noche, ~90 min) — #32 diseñado y decidido; código escrito, SIN probar ni commit

### Estado de los 5 grupos (sin cambios de fondo respecto a la V11)
| Grupo | Estado |
|---|---|
| 1 — Captura operativa | ✅ COMPLETO. **#32 en curso** aquí: el código de compras vive en el Grupo 1 |
| 2 — Inventario | ✅ COMPLETO |
| 3 — Integraciones | ✅ CERRADO (quedan #14 y #20 para hosting) |
| 4 — Costeo | ✅ COMPLETO |
| 5 — Financiero | 🔶 ~90%: faltan F6/F7 (#23/#24), #26, #27 |

### Commits y archivos
- **Bitácora V11: ya estaba subida** en el commit **aa770dd** ("Bitacora V11: #7 cerrado, m006 cafe, #30 resuelto,
  #32 nuevo"), `HEAD -> main, origin/main`. Confirmado con `git --no-pager log --oneline -3 -- <archivo>`.
- **Último commit: aa770dd** (bitácora). Último commit de código: **b3d666f** (m006). **Drive: versión 94 (sin tocar).**
- **SIN commit (a propósito, hasta probarlos)**, ya copiados al proyecto y verificados con Get-ChildItem en el destino:
  - `microservicios/grupo1_captura_operativa/ingreso_inventario.py` — **16941** bytes (antes 12543)
  - `microservicios/grupo1_captura_operativa/probar_ingreso_inventario.py` — **5934** bytes (antes 3096)
  - `git status` los mostrará como "modified" (normal).
- Respaldos en Escritorio\download: `ingreso_inventario_ANTES_32.py` (**12543**) y
  `probar_ingreso_inventario_ANTES_32.py` (**3096**), fechas originales 23-sep.
- Corrección de ruta: el código que registra compras está en **grupo1_captura_operativa** (no en grupo2_inventario,
  como decía el prompt V11). Función `escribir_fila_ingreso` (línea 173).

### 1) #32 — DECISIONES DEL USUARIO
- Aclaración del usuario (importante para la tesis): "el robot NUNCA escribe precios" significa que **nadie inventa
  ni digita un precio: el precio viene SIEMPRE de la TABLA CRUCE MAESTRA / MP**. El #32 agrega que **el precio ya
  registrado no debe cambiar**. Se descartó la opción de digitar el precio de la factura.
- Forma técnica: en Excel una fórmula viva siempre se recalcula; por eso la "conexión" con la TCM la hace el robot
  **al registrar** (lee el costo calculado de `wb_valores`) y escribe el resultado como **número** en E.
- **Decisión 2 = opción A:** congelar **solo compras** (columna E). La producción (I "Precio Unitario2", J) queda como
  fórmula; `hornadas.py` no se toca.
- **Decisión 3 = opción A:** si la TCM no tiene un costo > 0 para el producto, **se rechaza la factura completa sin
  escribir** (excepción nueva `CostoNoDisponible`), igual que con un nombre mal escrito.

### 2) Evidencia del Excel v94 (archivo de Drive, 247,159 bytes)
- TablaIngresos A3:L37 = 34 filas: **18 compras** (filas 4-20 y 37 = Natilla F004) y **16 hornadas** (21-36).
  (La V11 decía "34 compras": 34 es el total de filas.)
- E4 = fórmula viva idéntica en las 34 filas (TCM "Costo Unitario (Robot)" y si no MP "Costo Unitario (Scraping)",
  por NOMBRE). Ningún costo de compra es 0 (el menor: levadura 735).
- Quién lee Ingreso de Inventario: **Control de Inventario** C, D, G, H (nombres y cantidades) y **Flujo caja diario**
  A, **F**, **L** (fecha, total de compra, categoría). **Nadie lee E, I ni J directamente.**
- Congelar E con su valor actual **no cambia ningún número** (2-mar sigue −359,333.13; total 969,522.97).
- **Hallazgo:** el código escribía en L la `FORMULA_CATEGORIA` **vieja** (MP primero) mientras el Excel ya tiene la de
  la m006 en L4:L37 → una compra nueva de café habría salido "Materia Prima". Corregido en el código nuevo.
- Quién llama a `escribir_fila_ingreso` (git grep): `ingreso_inventario.py` (237), `hornadas.py` (183, 230, solo
  producción), `m009_inventario_prueba.py` (177 producción; 184 compra Natilla, ya aplicada).

### 3) Código nuevo — Partes A y B (escrito, verificado con simulación contra el v94; FALTA la prueba en la PC)
- `leer_costos_tcm(wb_valores)`: costo por NOMBRE, prioridad TablaProductos y luego MP (igual que la fórmula vieja).
- `_costo_congelado(nombre, costos)`: exige número > 0; si no, `CostoNoDisponible`.
- `registrar_ingresos`: revisa el costo de TODAS las compras de la factura antes de escribir; pasa
  `costo_unitario` a `escribir_fila_ingreso`. Resultado y evento incluyen "costo congelado 17,510.00".
- `escribir_fila_ingreso(..., costo_unitario=None)`: parámetro **opcional al final** (no rompe hornadas ni m009). Compra
  con costo → E = número; sin costo (llamada vieja) → fórmula de antes (la m012 la congela). F = D×E sigue fórmula.
- `FORMULA_CATEGORIA` = texto exacto de la m006 (verificado igual a L4). Docstring actualizado sin borrar la idea original.
- Simulación: harina E=17,510, leche 1,050, café Cf01 2,075.5 (números); hornada igual que antes; costo 0/vacío rechazado.
- Prueba (`probar_ingreso_inventario.py`): pasos 6-8 nuevos (E es número = costo TCM y F = D×E; café Cf01 con costo por
  paquete y L "Compras Externas Netas"; costos 0, None, −5, "abc" rechazados). Debe terminar en
  **"RESULTADO: TODO OK (0 lineas ERROR)"**.

### 4) Hallazgo pendiente de decisión — café Ca01 vs Cf01
- El café tiene dos códigos: **Cf01** (TCM, paquete, 2,075.5) y **Ca01** (MP, caja de 10, 20,755). Registrar "Ca01,
  cantidad 1" pensando en una caja se cuenta como **1 paquete** a 2,075.5 (inventario en paquetes, Regla #32).
  Ya pasaba antes del #32; no empeora. Opción a evaluar: rechazar Ca01 en compras y sugerir "use Cf01, cantidad en
  paquetes (1 caja = 10)".

### 5) Parte C pendiente — `migraciones/m012_congelar_costo_compras.py`
- Revisa antes de escribir: exactamente **18 filas** de compra con la fórmula vieja de E y valor > 0; si no, se detiene.
- Cambia E por su valor actual en esas 18 filas (hornadas no se tocan); nota en DiccionarioDatos con el texto de la
  fórmula vieja **sin el "=" inicial**; evento MIGRACION; 2a corrida = "Ya estaba aplicada".
- Esperado en comparar_libros: **18 fórmulas cambiadas, 0 valores cambiados**; VAN/TIR/punto de equilibrio iguales.
- Adjuntar `migraciones\m006_cafe_venta.py` (9134 bytes) como molde de estilo.
- Mapa V15: **Regla #33 nueva** ("El costo de una compra viene de la TCM y se congela al registrarla") y #32 resuelto.

### Pendientes (en este orden, próximo chat)
1. Correr la prueba local: `py -m microservicios.grupo1_captura_operativa.probar_ingreso_inventario` → TODO OK.
2. Construir y aplicar la **m012** con el protocolo completo (foto previa `maestro_prueba_ANTES_m012.xlsx`, --local,
   2a corrida, comparar_libros, valores clave, Drive ×2, probar_conexion).
3. Commit (archivos por nombre): `ingreso_inventario.py`, `probar_ingreso_inventario.py`, `m012_...py`; push.
4. Decidir Ca01/Cf01 del café.
5. Mapa V15 (Regla #33, #32 resuelto, Cambios V15).
6. F6/F7 con el tutor (#23/#24); #27 romper_prueba.py; #29 textos; #26 notas de tesis; #4; #21.
7. Después (otro chat): bot de WhatsApp real y bilingüe + hosting (ver V11, pendiente 6).

### Tips aprendidos hoy
- "working tree clean" con un archivo presente en la carpeta = Git ya lo tiene guardado. Para saber en qué commit:
  `git --no-pager log --oneline -3 -- NOMBRE_DEL_ARCHIVO`.
- `git --no-pager log --oneline --stat -2` muestra los últimos commits y QUÉ archivos incluyó cada uno.
- `git --no-pager grep -n "texto"` encuentra en qué archivo y línea está una función (así se corrigió la ruta).
- Adjuntar al chat: el nombre con comillas se escribe en la **ventana "Abrir" de Windows**, no en el chat.
- Verificar SIEMPRE en el destino con Get-ChildItem después de copiar: así se detectó que el paso 4b no se había corrido.
- El Explorador redondea a KB; la Terminal da los bytes exactos (así se identifican los archivos).

---

# CONTENIDO V11 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V11

*Última actualización: 26-sep-2026 (noche, ~9:45 p. m. CR). Reemplaza a la V10. La V10 completa sigue
abajo SIN CAMBIOS (sección "CONTENIDO V10", que a su vez trae la V9 y la V8); lo nuevo está en este bloque
"ACTUALIZACIÓN V11". Subir este archivo y el libro de mapas V14 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: los textos escritos esta noche en el código (pedidos_digitales.py, m006) y en DiccionarioDatos
dicen "27-sep-2026" (fecha UTC del servidor); en Costa Rica era el 26-sep-2026 por la noche. Es el mismo día.*

---

## ACTUALIZACIÓN V11 (26-sep-2026, noche) — #7 cerrado (Grupo 1 al 100%), m006 café, #30 resuelto, #32 nuevo

**Sesión de ~2 h de trabajo (4:12 a 6:05 p. m. CR) + cierre después de la cena (~9:00 a 9:45 p. m.).**

### Estado de los 5 grupos de microservicios (26-sep-2026, noche)
| Grupo | Estado |
|---|---|
| 1 — Captura operativa | ✅ COMPLETO AL 100% (#7 cerrado hoy) |
| 2 — Inventario | ✅ COMPLETO (nuevo pendiente de diseño: #32, congelar el costo de las compras) |
| 3 — Integraciones | ✅ CERRADO (quedan #14 y #20 para la fase de hosting) |
| 4 — Costeo | ✅ COMPLETO |
| 5 — Financiero | 🔶 ~90%: m006 hecha hoy; #30 resuelto. Faltan F6/F7 con el tutor (#23/#24), #26 notas de tesis, #27 romper_prueba.py |
**4 de 5 grupos completos.** Después: bot de WhatsApp real + hosting.

### Commits nuevos (agregar a la lista de la sección 7)
- Commit de la bitácora V10: ya estaba hecho y subido (**abf78b0**; confirmado con `git status` "up to date").
- **f37dc70** — #7: `microservicios/grupo1_captura_operativa/pedidos_digitales.py` (28132 bytes) +
  `probar_pedidos_digitales.py` (9092 bytes). 2 archivos, 280 líneas agregadas, 38 quitadas.
- **b3d666f** — `migraciones/m006_cafe_venta.py` (9134 bytes, archivo nuevo, 194 líneas).
- Pendiente: commit de esta bitácora V11 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V11.md`).
- **Drive: versión 94** (m006). Al cierre: `git status` = "nothing to commit, working tree clean".
- Respaldos en Escritorio\download: `pedidos_digitales_ANTES_7.py` (19342) y `probar_pedidos_digitales_ANTES_7.py`
  (5297). Foto previa: `datos_prueba\maestro_prueba_ANTES_m006.xlsx` (248126).

### 1) #7 CERRADO — anular la comanda completa (Regla #30) — commit f37dc70
- `anular_pedido_digital(id_pedido, motivo, momento=None, origen="SaaS")`, solo la administración; el bot nunca anula.
  - Interface digital: fila **WA-xxx-A** con SUBTOTAL, DESCUENTO, ENVÍO y TOTAL en **negativo**, ESTADO
    **"CONFIRMADO"**, la **fecha del pedido original**, DETALLE_COMANDA "ANULADO: motivo (anula WA-xxx)", mismo
    METODO_PAGO. Los montos se leen de `wb_valores` (valores ya calculados).
  - Detalle comanda digital: una fila de rastro WA-xxx-A por cada línea CONFIRMADA (mismo producto y cantidad,
    SUBTOTAL_LINEA negativo, ESTADO **"ANULADO"**). No devuelve inventario: Control de Inventario E solo suma
    CONFIRMADO y la línea original sigue CONFIRMADO. Tabla Combos C88 tampoco cuenta el rastro.
  - Evento **ANULACION_PEDIDO** en RegistroEventos.
  - Rechazos (excepción `AnulacionNoPermitida`, sin escribir): pago **SINPE** (pago en firme), pedido inexistente
    (`LookupError`), ya anulado, fila -A, motivo vacío, ESTADO distinto de CONFIRMADO.
- `anular_linea_pedido_digital` se conserva y **siempre** responde con un rechazo.
- **Forma de pago OBLIGATORIA (Regla #31)**: `registrar_pedido_digital` exige `metodo_pago` = "Efectivo" o
  "SINPE Móvil" (acepta "sinpe", "SINPE Movil", etc.; escribe siempre la forma exacta; excepción
  `MetodoPagoInvalido`). El resumen devuelve **`metodo_pago`** para que la comanda diga cómo paga el cliente.
  Pedidos viejos sin forma de pago = se tratan como Efectivo (se pueden anular).
- Revisado antes: `git --no-pager grep` → ningún otro código usa `registrar_pedido_digital` ni
  `anular_linea_pedido_digital` (solo las bitácoras y los dos archivos).
- **Prueba en la PC: TODO OK** (0 líneas ERROR). WA-119 (Cañas, Efectivo, ₡7,142.15 = 6,592.15 + envío 550) anulado
  con WA-119-A (Interface fila 28, −7,142.15 / envío −550; Detalle filas 41-43). WA-121 SINPE rechazado. Flujo
  caja: ventas digitales +1,685.38 = esperada (2 retiros de 842.69); envíos 0 = esperado.
- Por qué funciona sin tocar el Excel: Flujo caja diario C y K = SUMIFS por fecha (B) y ESTADO "CONFIRMADO".

### 2) Decisiones del usuario sobre pedidos digitales (para la fase del bot)
- El cliente **decide y elige en el bot** "Efectivo" o "SINPE" (pago electrónico); la forma de pago llega **junto con
  la comanda**. Las líneas solo se quitan en la conversación **antes de "Confirmar pedido"** (cuando entra al Excel).
- **SINPE = pago en firme** (no se anula). Efectivo se cobra al entregar; si el cliente no lo quiere, la
  administración puede anular la comanda completa. **Al entregar el pedido se muestra el recibo** (fase del bot /
  reparto). La comanda debe decir cómo paga (p. ej. "Pago: Efectivo (cobrar al entregar)").
- **La comanda es UN SOLO RENGLÓN** (Interface digital, DETALLE_COMANDA: una fila por pedido, productos separados por
  coma). El bot entrega un **JSON** (confirmado por el usuario), ejemplo:
  `order_id`, `customer {phone, name}`, `items [{product, qty, unit_price, subtotal}]`, `total`.
  Traducción: order_id → id_mensaje; customer.name → cliente; items → lineas (product/qty → producto/cantidad);
  total → total_cotizado. **Aprobado**: el traductor JSON → registrar_pedido_digital se construye en la fase del bot.
  Cuidar: nombres de producto = código o nombre exacto de la TCM (Regla #26; "Pan Baguette Relleno" se rechazaría);
  el robot NO usa unit_price/subtotal del bot (Regla #4); faltan en el ejemplo forma de pago, ubicación/km y envío;
  teléfono: decidir si se guarda (#33).

### 3) Sugerencia de arquitectura SaaS del usuario (solo una sugerencia)
- Archivo `saas_architecture_document.md`: 5 servicios = 5 grupos + módulo 0 de configuración, API Gateway,
  FastAPI, PostgreSQL/Redis, React/Next.js, JWT + OAuth2 multi-tenant, 4 dashboards. El usuario: "lo de SaaS se ve en
  el momento".
- Encaja: 5 servicios = 5 grupos; FastAPI como puente; la mayoría de endpoints ya tienen función del robot; la idea de
  una **fila de procesamiento** (escrituras de una en una) resuelve que bot y SaaS no se pisen.
- Choca (a decidir en su momento): PostgreSQL como segunda "verdad" (la fuente es el Excel en Drive, Reglas #12/#27);
  multi-tenant (basta un negocio con roles); combos y mermas ya resueltos/no existen; Módulo 4 necesita PUT (costeo.py
  escribe); "Reglas Globales" vive en el libro de mapas, no en el Excel.
- Endpoint futuro: `POST /api/v1/orders/{id_pedido}/cancel` solo rol administración → anular_pedido_digital.
- Pregunta sugerida para el tutor: ¿el entregable incluye las pantallas del SaaS o basta el robot + una API sencilla?

### 4) m006_cafe_venta — APLICADA (local y Drive v94), commit b3d666f — resuelve #25
- Decisión del usuario: el café es **COMPRA EXTERNA** (se revende tal cual; comprarlo en caja de 10 no lo hace
  materia prima). Que saliera "Materia Prima" era un **hueco** (Regla #32).
- A) TCM fila 15 (reserva dentro de TablaProductos C6:I15): C15 "Café Naranjo puro 250gr x 10 Unidades" (mismo nombre
  que Control A21), D15 = MP!$D$19/10 (**2,075.50**), E15 **Cf01**, F15 "Café", G15 = costo × 1.05 (**2,179.28**),
  H15 "Scraping Solera", I15 "Tercero". La fila 19 de la MP se conserva.
- B) Ventas despacho L13 = TCM C15; lista E5:E27 → $L$2:$L$13. C) Ingreso D19 1 → **10 paquetes** (F19 sigue 20,755).
  D) Ingreso L4:L37: TablaProductos primero → "Compras Externas Netas"; si no, MP → "Materia Prima" (solo cambia el
  café). E) DiccionarioDatos N56 + fila 305. F) evento.
- `precios_proveedores.py` NO toca D15 (Ca01 solo se busca en la MP col F; en la TCM solo L01, Q01, Ma01, CC01, Na01).
- La migración revisa todo antes de escribir (se detiene sin escribir si algo no coincide) y responde "Ya estaba
  aplicada" en la 2a corrida (probado en local y Drive).
- comparar_libros: 5,016 fórmulas iguales, 69 celdas, 8 valores en cascada, **0 errores**. Valores clave: Ingreso
  E19 20,755 → 2,075.5; L19 → Compras Externas Netas; Control F21 1 → 10, H21 REABASTECER → OK; Flujo 2-mar E5
  490,800 → 470,045 y F5 61,200 → 81,955; H5, I5, N11, N12, flujo anual N25 16,522,417.18, TIR VAR C10
  10,950,907.94 y PE B29 46,054 **sin cambio**. probar_conexion: 8 productos, Cf01 ₡2,179.275.

### 5) #30 RESUELTO — el 2-mar negativo (opción A, sin tocar el Excel)
- I5 = −359,333.13: el día 1 se compra el inventario inicial para producir (egresos 552,000 vs ventas 192,666.87).
- 'Flujo caja diario' **lee** Ingreso de Inventario (compras; costo por fórmula desde TCM/MP), Ventas despacho e
  Interface digital (precios congelados al vender). **Ninguna hoja lo lee a él**: es control operativo; no afecta
  flujo anual, flujo del proyecto, TIR ni PE. El Historico Diario (hornadas) no mueve dinero.
- Ese dinero está en el **capital de trabajo** de 'flujo caja proyecto' (D24 = B45 = 4,074,407.94; B37 Insumos
  iniciales = 1,290,690.69 > 552,000). El usuario: "debe haber plata en caja para pagar la compra; es parte del
  capital de trabajo". Opción C (saldo inicial de caja) descartada: el capital de trabajo ya cumple ese papel.

### 6) #32 NUEVO — el costo de las compras es fórmula viva (decidir en el próximo chat)
- Ingreso de Inventario E = INDEX/MATCH a TCM/MP **sin fecha**: si cambia un precio, **todas las compras pasadas se
  recalculan**. Ensayo (harina +10%, MP D11 17,510 → 19,261): Ingreso F11 del 2-mar 280,160 → **308,176**; Flujo E5
  490,800 → 518,816; H5 552,000 → 580,016; I5 −359,333.13 → −387,349.13.
- Las ventas no tienen el problema (el robot escribe el precio como número fijo). Propuesta: el robot escribe el
  costo como número fijo al registrar la compra + migración que congele las 34 compras. Estimado: 60-90 min.

### Libro de mapas V14 (LIBRO_MAPA_DE_CELDAS_PFV14.xlsx)
- Pestaña nueva "Cambios V14" (11 filas); Regla #30 actualizada; **Reglas nuevas #31** (forma de pago obligatoria)
  y **#32** (compra externa vs materia prima); Decisiones Pendientes: #7, #25, #30 resueltos; nuevos **#32** y **#33**;
  bloques "ACTUALIZACIÓN V14" en 7 hojas; "|| V14" en Detalle E4/E6/H4 y Diseño del robot (2) C20/D17.
- Verificado: 6,719 celdas iguales, 9 con texto agregado "|| V14", 20 nuevas + pestaña nueva, **0 borradas**,
  0 cambiadas, 47 comentarios intactos.

### Pendientes (en este orden, próximo chat)
1. Commit de esta bitácora V11 si no quedó hecho.
2. **#32** decidir y construir: congelar el costo de las compras (Grupo 2 + migración). Adjuntar el código que
   registra compras (`microservicios\grupo2_inventario\...` ingresos) y el Excel maestro recién descargado.
3. **F6 / F7** con el tutor (#23 tasa de impuesto, #24 financiamiento) — el usuario trae la decisión.
4. #27 revisar `romper_prueba.py`; #29 textos cosméticos (probar_hornadas.py líneas 51-52; notas viejas del
   DiccionarioDatos); #26 notas de metodología en la tesis (incluir #30); #4 dos líneas idénticas en una factura.
5. #21 (usuario): clave de Gemini en el libro de mapas, NODO INTEGRAL E52.
6. Después (otro chat): bot de WhatsApp real y bilingüe (#16, #28, #33: traductor JSON, recibo al entregar, forma de
   pago en la comanda) y hosting (#14, #20; Render + GitHub Actions a las 7:00; respaldo diario del Excel; endpoint
   de anulación solo administración; fila de escrituras).

### Tips aprendidos hoy
- Si el navegador guarda un archivo en la raíz de ROBOT PANADERO (u otra carpeta), buscarlo con
  `Get-ChildItem "carpeta1", "carpeta2" -Filter "*nombre*" | Select-Object FullName, Length` y **moverlo** con
  `Move-Item` (con Test-Path antes) para no dejar copias sueltas que Git mostraría como nuevas.
- Para reemplazar un archivo del proyecto: 1) respaldo del viejo en download con otro nombre (`_ANTES_xx`),
  2) `Copy-Item ... -Force`, 3) `Get-Item` del DESTINO con el tamaño nuevo.
- Leer valores clave antes → después sin abrir Excel: `py -c "import openpyxl ..."` con `data_only=True` sobre la foto
  previa y la copia (solo lee).
- El nombre del Excel en Drive no cambia con las versiones ("v68" en el nombre); la versión real la da probar_conexion.
- `git status` con "nothing to commit, working tree clean" + "up to date with 'origin/main'" = todo guardado en
  GitHub.

---

# CONTENIDO V10 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V10

*Última actualización: 26-sep-2026 (tarde, ~3 p. m. CR). Reemplaza a la V9. La V9 completa sigue
abajo SIN CAMBIOS (sección "CONTENIDO V9", que a su vez trae la V8); lo nuevo está en este bloque
"ACTUALIZACIÓN V10". Subir este archivo y el libro de mapas V13 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

---

## ACTUALIZACIÓN V10 (26-sep-2026) — Grupo 3 cerrado; Grupo 5 casi completo (m008 a m011)

**Sesión de ~6 h (9:00 a. m. a 3:00 p. m. CR).**

### Estado de los 5 grupos de microservicios (26-sep-2026)
| Grupo | Estado |
|---|---|
| 1 — Captura operativa (ventas despacho, pedidos digitales, ingreso de inventario, catálogo) | ✅ COMPLETO (ajuste de código pendiente: #7, anular comanda completa) |
| 2 — Inventario (hornadas, consumo de materia prima) | ✅ COMPLETO |
| 3 — Integraciones (NODO INTEGRAL, Google Maps, precios de proveedores, competencia Musi) | ✅ CERRADO HOY (quedan #14 y #20 para la fase de hosting) |
| 4 — Costeo (m005 + costeo.py) | ✅ COMPLETO (25-sep) |
| 5 — Financiero | 🔶 ~80%: m007 (25-sep) + m009, m010, m011 (hoy). Faltan #7 (código), m006 café (#25), F6/F7 con el tutor (#23/#24), notas de tesis (#26), romper_prueba.py (#27) |
**4 de 5 grupos completos.** Después: bot de WhatsApp real + hosting.

### Commits nuevos (agregar a la lista de la sección 7)
- Commit de la bitácora V9: hecho al cierre del chat anterior (**cf6f9c6**, confirmado por el usuario).
- **191ffe5** — `migraciones/m008_nombres_unicos.py` + comentarios de `microservicios/grupo2_inventario/hornadas.py` (13289 bytes).
- **751f6dc** — `migraciones/m009_inventario_prueba.py` (11586 bytes).
- **0484ad8** — `migraciones/m010_descuento_combos.py` (6021 bytes).
- **b6f2b69** — `migraciones/m011_ventas_digitales.py` (12995 bytes).
- Pendiente: commit de esta bitácora V10 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V10.md`).
- `m008_nombres_unicos.py` = 10685 bytes. Drive: v79 (m008), v81 (Musi), v85 (m009), v88 (m010), **v91** (m011).
  Drive NO suma un número fijo de versiones por escritura (no predecirlo).

### 1) Grupo 3 CERRADO — lectura real de Musi
- 1er intento (--local): FALLO por lectura incompleta **7 de 9** (Gemini saturado, HTTP 503 en los 3 modelos
  flash; 503 = servidor ocupado, distinto de 429 = sin cupo). Celdas intactas (Regla #19), evento registrado.
- 2o intento: **9 de 9 → SIN_PROMO**. El cache hizo que solo se leyeran las 2 imágenes que faltaban; la imagen 9
  (Café NESCAFÉ) tenía fecha impresa "hasta 2025-06-30" y Gemini la usó. Escrito en local y Drive (sin --local,
  todo desde cache, en segundos): NODO INTEGRAL C21 y Tabla Combos D11 = "Sin promocion vigente (26-09-2026)",
  E11 = 0. probar_conexion TODO BIEN.
- Queda para hosting: #14 (tarea de las 7:00 + clima/tipo de cambio) y #20 (cache en `CACHE_COMPETENCIA.xlsx`:
  falta el código; sin él, en la nube Gemini releería las 9 imágenes cada día y agotaría el cupo).
- Pendiente del usuario: #21 pegar la clave de Gemini en el libro de mapas, NODO INTEGRAL E52 (nunca en el chat).

### 2) Principio del usuario para el Grupo 5 (Regla #27)
Es un **estudio académico**: ningún dato es venta real; cada hoja se creó por partes con valores inventados y
hay huecos que ajustar. Se trata de **cuadrar un sistema completo con datos supuestos**. Las rebajas de inventario
las hace SIEMPRE el Excel con sus fórmulas; lo digitado son valores según el mapa. El supuesto del modelo
financiero (46,054 unid/año por producto, 'Punto de Equilibrio' B29/B30) no se toca.

### 3) m008_nombres_unicos — APLICADA (local y Drive v79), commit 191ffe5
- Decisión del usuario (Regla #26, REEMPLAZA la del Grupo 2): **un solo nombre por producto** en todo el
  documento, igual a la TABLA CRUCE MAESTRA: **"Baguette"** (C8) y **"CupCake"** (C9). Externos ya estaban
  alineados.
- Cambios: A 112 celdas de datos (Ventas despacho E, Interface F, Detalle C, Historico C), B 2 celdas
  ('costo …'!A6), C 120 fórmulas de 'flujo caja anual' (solo el texto entre comillas), D 24 etiquetas.
  No se renombran las hojas ni se reescribe DiccionarioDatos/RegistroEventos.
- comparar_libros: 269 celdas, 5 valores (Control de Inventario D6/E6/F6/H6 del CupCake + costo cupcakes I30
  = LARGO(A6) 8→7), 0 errores. VAN/TIR/PE/flujo anual idénticos (N25 16,522,417.18; TIR VAR C10 10,950,907.94).
- hornadas.py lee el nombre del Historico en A6 (sin cambios de lógica; solo 3 comentarios).

### 4) m009_inventario_prueba — APLICADA (local y Drive v85), commit 751f6dc — resuelve #18
- Semana de prueba = **2 al 8 de marzo de 2026**. El Historico ya tenía 16 hornadas de prueba
  (14 del 2-8 mar = 1,400 Baguette + 1,400 CupCake; 2 de mayo = 100 + 30; total 1,500 / 1,430), que cubren
  despacho (779 / 841) + digital de prueba (475 / 465).
- A) fechas A4:A9 → 2,2,4,5,6,7 mar; A10:A20 → 2 mar. B) compras justas: harina 16, levadura 19, mejorante 7,
  bolsas 15, burbuja 6, margarina 2, azúcar 1, sal 1, manteca 1, café 1, mermelada 1, leche D7 10→13 (33 L).
  C) producción vieja H4:H9 → 0 (relleno). D) 16 filas de producción + 96 de consumo (TablaConsumoMP), escritas
  con las funciones REALES del robot (`escribir_fila_ingreso`, `_escribir_consumo`). E) compra de Natilla F004 = 15
  (se vendían 12 y nunca se había comprado).
- Hallazgo: `escribir_fila_ingreso` solo escribe el Total del lado usado (F compra / J producción) para evitar
  vacío × vacío = #VALUE!; exige `es_produccion` (sin valor por defecto).
- comparar_libros: 1,063 celdas, 349 valores, 0 errores. Resultado: ninguna existencia negativa (Baguette 246,
  CupCake 124); materia prima de prueba 490,800; flujo neto de prueba −1,754,367 → +406,803.
- Nota: 'Control de Inventario' A6 dice "CupCake"; con la m008 D6/E6 ya encuentran las ventas.

### 5) m010_descuento_combos — APLICADA (local y Drive v88), commit 0484ad8 — resuelve #19
- Decisión del usuario (Regla #28): si sube el externo (queso, natilla…), **el combo sube aunque tenga
  descuento** (no se compensa); el **descuento lo digita la administración** en 'Tabla Combos Dinamicos' L.
- Hallazgo: L existía pero ninguna fórmula la leía (0.9 fijo). E5:E8: `0.9` → `(1-IF($L>1,$L/100,$L))`
  (acepta 10 o 10%; L vacía = sin descuento). El descuento sigue sobre el pan/cupcake (diseño de la tabla).
- Revisión de dependencias: ninguna hoja lee E (Ventas despacho L9:L12 lee C; Control I lee
  TablaComposicionCombos). Código: catalogo.py lee el VALOR de E y L; pedidos_digitales.py líneas 167-170 usa L
  solo como trazabilidad (DESCUENTO_ASIGNADO), descuento en colones = 0 → **sin doble descuento**.
- comparar_libros: 15 celdas, **0 valores cambiados**, 0 errores (precios 3377.76 / 1405.19 / 2398.84 / 4690.26).
  Probado aquí: L=5 → Combo 1 = 3504.17; L=5% igual; L vacía → sin descuento.

### 6) m011_ventas_digitales — APLICADA (local y Drive v91), commit b6f2b69 — resuelve #15
- Decisión del usuario (Regla #29): **una cosa es la venta digital y otra el cobro del envío**. En 'Interface
  digital' TOTAL_A_COBRAR = comanda + delivery es CORRECTO y no cambia; 'Detalle comanda digital' NO lleva envío.
- Hallazgo: WA-102..WA-118 (semana simulada creada el 24-25 sep con semilla 42, base 184 unid/día/producto con
  multiplicadores por día) no tenían TOTAL_A_COBRAR, METODO_PAGO ni ESTADO → el flujo solo veía 5,500.
- A) 51 celdas: J = subtotal − descuento + envío, K = Efectivo / SINPE Móvil alternados (decisión del usuario:
  mezclados, más cerca de la realidad), M = CONFIRMADO. B) WA-101 (2-abr = Jueves Santo): "10 Baguette",
  8,426.90 − 421.35 (feriado 5%) + 500 = 8,505.55; Detalle fila 4: 8,005.55, FERIADO, 0.05.
  C) 'Flujo caja diario': C = SUMIFS(G) − SUMIFS(H); columna NUEVA K "Envíos cobrados" = SUMIFS(I); D = B+C+K;
  tabla A4:J369 → A4:K369; resumen movido K3:L13 → M3:N13 + fila M14:N14. D) nota en DiccionarioDatos N28.
- comparar_libros: 1,197 celdas, 217 valores, 0 errores. Semana de prueba: ventas digitales 559,219.90, envíos
  9,000, total ingresos 1,560,172.97, flujo neto **969,522.97**. VAN/TIR/PE/inventario idénticos.
  El 2-mar sale negativo (−359,333) porque se compra la materia prima de toda la semana (#30).

### 7) #7 — anulación de pedido digital: DISEÑO ACORDADO (código pendiente) — Regla #30
- El bot pregunta "Confirmar pedido" y solo con el "Sí" registra (ya existe: cotizar_envio → cotizar_pedido →
  registrar_pedido_digital con total_cotizado). Ese "Sí" es el compromiso: **si pide 3 artículos paga 3**.
- **No se anulan líneas sueltas.** Si al llegar el cliente no lo quiere, el cliente o el repartidor llaman y la
  **administración decide**; si se anula, es la **comanda completa**. Un pedido pagado es firme (lo maneja la
  administración). El bot nunca anula.
- Diseño: `anular_pedido_digital(id_pedido, motivo)` (usa la administración desde el SaaS):
  Interface digital = fila negativa del TOTAL (comanda + envío; el envío es parte del pedido) con ESTADO
  "CONFIRMADO" y la **fecha del pedido original**, ID p. ej. `WA-119-A`, detalle "ANULADO: motivo";
  Detalle comanda digital = filas de rastro con ESTADO **"ANULADO"** que NO devuelven inventario (los alimentos
  que salen de la panadería no se reciben de vuelta); evento en RegistroEventos. La función vieja
  `anular_linea_pedido_digital` queda respondiendo con un rechazo (no se borra).
- **Confirmar con el usuario al iniciar**: cómo trata el robot un pedido con METODO_PAGO "SINPE Móvil"
  (propuesta: rechazar la anulación porque el pago es firme).
- Para construir: adjuntar `probar_pedidos_digitales.py` (hoy espera el comportamiento viejo "ANULADO" por línea).
- Para la fase del bot (#28): mensaje de confirmación sugerido: "Al confirmar, tu pedido queda en firme. Por
  higiene, los alimentos que salen de la panadería no se reciben de vuelta."

### Libro de mapas V13 (LIBRO_MAPA_DE_CELDAS_PFV13.xlsx)
- Pestaña nueva "Cambios V13" (12 filas); Reglas #26 a #30; Decisiones Pendientes: #13, #15, #18, #19, #22
  resueltos, #7 diseño acordado, nuevos #28 a #31; bloques "ACTUALIZACIÓN V13" en 13 hojas.
- Verificado: 6,598 celdas iguales, 6 con texto agregado "|| V13", 44 nuevas + pestaña nueva, 0 borradas,
  47 comentarios intactos.

### Pendientes (en este orden, próximo chat)
1. **#7** anular comanda completa (código en pedidos_digitales.py + probar_pedidos_digitales.py). Confirmar SINPE.
2. **m006** café a la venta por paquete de 250 g (#25).
3. **F6 / F7** con el tutor (#23 tasa de impuesto, #24 financiamiento).
4. #26 notas de metodología en la tesis; #27 revisar `romper_prueba.py`; #29 textos cosméticos
   (probar_hornadas.py líneas 51-52); #30 decidir si se explica el 2-mar negativo; #4 dos líneas idénticas en
   una factura.
5. #21 (usuario): clave de Gemini en el libro de mapas, NODO INTEGRAL E52.
6. Después (otro chat): bot de WhatsApp real y bilingüe (#16, #28) y hosting (#14, #20; Render + GitHub Actions
   a las 7:00; respaldo diario del Excel).

### Tips aprendidos hoy
- Si la Terminal muestra texto con `:` o `(END)` abajo, es el **visor (less)** de git: se sale con **q**. Si
  pregunta "Log file … (press RETURN)", NO dar Enter: **Ctrl+C** y luego **q**. Para evitarlo:
  `git --no-pager grep …` / `git --no-pager diff …`. (Ese día quedó un archivo basura "t-Path…"; se borró con
  `Get-ChildItem -File -Filter "t-Path*" | Remove-Item`, ensayado antes con `-WhatIf`.)
- `>>` al inicio de la línea = comando incompleto (falta una comilla): **Ctrl+C** y escribirlo de nuevo.
- No copiar las líneas de EJEMPLO de lo que "deberías ver"; solo los recuadros de comandos.
- `git push` sin punto al final (`git push.` falla). Repetir `git commit` dice "nothing to commit" (inofensivo).
- Ctrl+P en VS Code abre el buscador de archivos (Esc lo cierra). Los atajos van a la parte donde se hizo el
  último clic: clic en la Terminal antes de escribir.
- Si el navegador guarda un archivo con "(1)", ya existía uno con ese nombre: identificar SIEMPRE por tamaño.
- Un comando `if … else { Copy-Item …; "COPIADO" }` puede imprimir COPIADO aunque falle: verificar con
  `Get-Item "destino"` (tamaño en la carpeta de destino).
- El aviso "LF will be replaced by CRLF" es normal en Windows.
- El Excel descargado en la mañana queda viejo después de cada migración: para revisar, descargar de nuevo de
  Drive (nunca abrir con Google Sheets).

---

# CONTENIDO V9 (sin cambios)

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
