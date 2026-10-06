# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V24

*Última actualización: 5-oct-2026 (~9:15 p. m. CR). Reemplaza a la V23. La V23 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V23", que a su vez trae la V22 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V24".
Subir este archivo y el libro de mapas V26 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.** (Las claves del SaaS están solo en el `.env`, en Render y en
el libro de mapas, Regla #15.)

*Nota de fechas: el chat fue el 5-oct-2026, de las 5:35 p. m. a ~9:15 p. m. (hora CR); es el TERCER chat del día (V22 la
mañana, V23 la tarde). Las horas de las pruebas salen de la Terminal del usuario, de 'RegistroEventos' y de los Logs de Render.*

---

## ACTUALIZACIÓN V24 (5-oct-2026, noche, ~3 h 40 min) — SaaS Grupo 4 (costeo) COMPLETO: hojas de costo con precio de referencia y proporciones, planilla, gasto operacional y valor activos, probado en Render; `costeo.py` recibe `origen` (#78); nota de C8 (#96); CupCake corregido en los tres Excel: panadero y ayudante a 20 % y precio de referencia 391.92

### Estado general
- **Grupo 4 del SaaS COMPLETO** en dos paquetes, los dos en GitHub y probados en Render: **621dfae** (paquete 1) y
  **ad390ac** (paquete 2). Menú nuevo "4 Costeo" con cinco pantallas: Costo pan baguette, Costo cupcakes, Planilla, Gasto
  operacional y Valor activos.
- **Función nueva `actualizar_proporcion`** (decisión del usuario): la administración define desde el SaaS qué parte del
  sueldo de panadero y ayudante, y qué parte del agua y de la electricidad, se carga a cada producto.
- **CupCake corregido por el usuario en la copia DEMO, la copia local y el Excel MAESTRO:** panadero y ayudante de 15 % a
  **20 %** ("fue error mío") y precio de referencia de 343.16 a **391.92** ("343.16 no deja ganancia"). Sin migración: son
  valores de configuración, escritos con las funciones de `costeo.py`, cada uno con su evento.
- **Permisos (decisión del usuario):** hojas de costo, planilla y gasto operacional son SOLO para administración; Valor
  activos es lectura para los dos usuarios.
- **#96 RESUELTO** (la lista de C8 no lleva "Bajar"; la nota dice que es una prevista). **#97 aclarado. #99 cerrado por
  diseño** (el robot no exige un mínimo de hornada).
- **#57:** sin dato nuevo (mismo día 5-oct). El segundo día de los dos horarios se revisa el 6-oct.
- El chat se cierra aquí por decisión del usuario. El Grupo 5 (#66 a #69) y el Registro de eventos van en el chat nuevo.
- Libro de mapas **V26**: pestaña "Cambios V26"; nuevos **#99 a #105**.

### Commits de este chat (todos con push)
- **1274aa1** — bitácora V23 (hecho al cierre del chat de la tarde; es la marca del `repo_actual.zip` adjunto).
- **621dfae** — "SaaS Grupo 4 paquete 1: hojas de costo, precio de referencia y origen SaaS en costeo (#78, #96)":
  8 files changed, 553 insertions, 22 deletions; 2 archivos nuevos. Push `1274aa1..621dfae`. 74 commits en GitHub.
- **ad390ac** — "SaaS Grupo 4 paquete 2: proporciones de costo, planilla, gasto operacional y valor activos; costeo solo
  para administracion": 10 files changed, 1027 insertions, 22 deletions; 3 archivos nuevos. Push `621dfae..ad390ac`.
- **Último commit de código: ad390ac.** Pendiente al cierre: commit de esta bitácora V24 (solo este archivo).

### 1) Revisión de los adjuntos al inicio
- El prompt V23 llegó como archivo adjunto, sin texto en el mensaje; se leyó desde el archivo.
- `repo_actual.zip`: marca **1274aa1**, 1,095,371 bytes, 132 archivos, sin `.env` ni credenciales. Bitácora V23 250,625;
  mapa V25 307,536.
- Excel maestro: **253,707 bytes**, 25 hojas, 0 errores, 46 eventos; **EVT-000046 MIGRACION m015** (5-oct 14:56:47);
  'TABLA CRUCE MAESTRA'!G12 = 1,050 (m015 comprobada). `TablaComposicionCombos` A15:H23.
- Copia DEMO: **256,073 bytes**, 25 hojas, 0 errores, 62 eventos; **EVT-000057 MIGRACION m015** (15:00:26); G12 = 1,050.
  Los tres COMPOSICION_COMBO son EVT-000060 (16:52:31), EVT-000061 (16:57:35) y EVT-000062 (17:01:00); tabla A15:H24 como
  se esperaba (fila 16 D = 3, G = 5, H = 15; fila 17 D = 1, G = 5, H = 0; fila 24 historia de la leche, H = 5).
- **#97 aclarado:** los dos CONFIG_COMBO son EVT-000058 (16:47:04) y EVT-000059 (16:51:52); solo cambiaron el mensaje K5
  del Combo 1 ("3 Baguette" → "4 Baguette" y de vuelta). Efecto neto cero.
- Capturas de GitHub Actions (5:48 p. m. y 7:39 p. m.): las mismas 10 corridas (#9 a las 00:43 y #10 a las 06:35).

### 2) Grupo 4 — evidencia de celdas (mapa y Excel coinciden en lo digitado a mano)
- 'costo pan baguette'!G28 = 842.69 y 'costo cupcakes'!G27 = 343.16 (precio de referencia, Regla #9); 'planilla' B y C
  (filas 7 a 12; plaza nueva 13 a 32); 'gasto operacional' filas 6, 7, 8, 11, 12, 14, 15, 17 y 18 (B a M); 'valor activos '
  A, B, C y E (filas 3 a 33). G4, la fecha, la hora y G2 de las hojas de costo ya los escribe la pantalla Hornadas (Grupo 2).
- `costeo.py` ya tenía las seis funciones desde el 25-sep (14 casos en `probar_costeo.py`).
- Hallazgos presentados: (1) las seis funciones no recibían `origen` (#78); (2) y (3) dos textos viejos del mapa
  'costo cupcakes' (MONTH($C$6) y leche a TCM G10; el Excel está correcto: #104 b); (4) las proporciones B23, B24, B31 y B32
  (0.8 en Baguette) y B23, B24, B30 y B31 (0.15 y 0.2 en CupCake) estaban escritas como número y el mapa no las marcaba.

### 3) Decisiones del usuario (en orden)
1. Alcance: **paquete 1 y luego paquete 2.** #96: la lista de C8 **NO lleva "Bajar"**.
2. **Las proporciones de las hojas de costo las define la administración desde el SaaS** ("hoy está esta proporción de
   consumo pero puede variar"); van en el paquete 2 con tres reglas: de 0 a 100 %; por concepto la suma de los dos productos
   no pasa de 100 %; si el costo queda sobre el precio se guarda con aviso (#100).
3. **Hornada mínima (#99):** "como sistema semi industrial tiene una cantidad mínima de horneada que cubre el costo"; y
   después: "el robot no exige ese mínimo de horneada porque es un tema administrativo de costos". Queda por diseño.
4. **CupCake a 20 %** (panadero y ayudante): "fue error mío, no dejar como pendiente" (#101).
5. **Permisos (#102):** "dejar solo a administración ver las planillas y las hojas de costos" y "Gasto operacional no
   visible para despacho"; el permiso de gasto operacional fue dentro del paquete 2, que todavía no estaba instalado.
6. **Precio de referencia de CupCake = 391.92:** "343.16 es un error porque no deja ganancia". Lo había guardado él mismo
   en la copia DEMO desde el SaaS (19:48); Claude lo señaló como guardado fuera del plan y el usuario confirmó su selección.
7. **Cerrar el chat** con el protocolo de cierre al terminar esos pasos.
- Sin decidir: ocultar a despacho el costo unitario de la pantalla Hornadas (#104 d); existencia mínima de Baguette en 5.

### 4) Lo construido
**Paquete 1 — `paquete_grupo4_costo.zip`, 41626 bytes, 8 archivos (tamaños en la PC):** `costeo.py` **25727** (antes 20594:
dato opcional `origen` en las seis funciones y lectura `leer_costeo`), `probar_costeo.py` **7303** (antes 5734: caso 15 con
origen; los eventos CONFIG_ se cuentan como nuevos), `saas_demo/grupo4.py` **8666** (nuevo), `plantillas/costo.html`
**7703** (nuevo), `app.py` **14862**, `__init__.py` **1192**, `plantillas/nodo_integral.html` **6940** (nota de C8) y
`probar_saas_demo.py` **69323** (sección P).

**Paquete 2 — `paquete_grupo4_parte2_v4.zip`, 54866 bytes, 10 archivos** (reemplazó, sin instalarse, a los de 54336, 54361 y
54761): `costeo.py` **39376** (`actualizar_proporcion`, `leer_planilla`, `leer_gastos`, `leer_activos`), `probar_costeo.py`
**10438** (casos 16 a 21), `grupo4.py` **20962**, `probar_saas_demo.py` **83664** (sección Q), `app.py` **15197**,
`__init__.py` **1323**, `plantillas/costo.html` **10690** y tres plantillas nuevas: `planilla.html` **6358**,
`gasto_operacional.html` **4211** y `valor_activos.html` **6172**.

Cómo trabaja lo nuevo de `costeo.py`:
- **`origen`**: sin él el evento dice "Robot Python" como siempre; el SaaS entrega "SaaS (administracion)".
- **`leer_costeo(wb)`**: las dos hojas de costo tal como las calculó el Excel (materiales, mano de obra e indirectos con su
  total, costo unitario, precio de referencia, precio calculado y utilidad). Ubica las filas por el texto de la columna A.
- **`actualizar_proporcion(producto, concepto, porcentaje, origen)`**: conceptos `panadero`, `ayudante`, `agua` y
  `electricidad`. Recibe el porcentaje (80) y guarda la fracción (0.8). Comprueba la celda por el texto de su fila, rechaza
  si la suma con el otro producto pasa de 100 % y, después del recálculo, lee el costo nuevo: si queda sobre el precio de
  referencia agrega un AVISO. Evento **CONFIG_PROPORCION_COSTO**. Uso manual:
  `py -m microservicios.grupo4_costeo.costeo --local proporcion P01 panadero 75`.
- Por qué la suma no pasa de 100 %: 'gasto operacional' fila 9 (Planilla) resta de la planilla lo que las hojas de costo
  ya cargaron a producción (m007, Regla #24).

Pantallas (grupo "4 Costeo"): las dos hojas de costo (cifras, formulario del precio de referencia, tabla de proporciones
y formulario de proporción, elementos del costo); Planilla (cambiar una plaza, agregar una plaza); Gasto operacional (13
rubros por mes; formulario para los 9 manuales); Valor activos (cambiar o dar de baja con cantidad 0, agregar).

### 5) Pruebas
- **Entorno de Claude** (copia DEMO adjunta): `probar_costeo` 35 de 35; `probar_saas_demo` TODO OK, **208 revisiones**
  (antes 162), también sobre una copia con CupCake en 20 %. La prueba atrapó un error de Claude: la sección P daba por
  hecho que Baguette tenía margen, pero la sección K deja una hornada de 10 (de ahí salió el #99).
- **PC del usuario:** paquete 1: 8 tamaños correctos, `probar_costeo` 23 de 23, `probar_saas_demo` TODO OK (A a P).
  Paquete 2: 10 tamaños correctos, `probar_costeo` **35 de 35**, `probar_saas_demo` **TODO OK (A a Q)**; copia local intacta.
- **Tropiezo sin daño:** al pegar el bloque del commit, su primer renglón se unió al comando de la prueba que estaba
  escrito sin ejecutar (`probar_saas_demogit add`); la prueba y el `git add` no corrieron y el commit dijo "no changes
  added". Se repitió en orden.
- **Render, paquete 1 (administración, copia DEMO):** precio de referencia de Baguette 842.69 → 850.00 (13.8 s), 850.00 →
  843.69 (11.6 s, guardado intermedio no planeado) y 843.69 → 842.69 (11.6 s). Queda en 842.69.
- **Render, paquete 2:** las dos proporciones de CupCake (sección 6); Gasto operacional con la fila Planilla en 340,000;
  la nota de C8 "Es una prevista..."; con el usuario despacho, el menú "4 Costeo" solo ofrece Valor activos.

### 6) CupCake: proporción a 20 % y precio de referencia 391.92 (#101)

| Archivo | Cómo | Panadero B24 | Ayudante B23 | Precio G27 |
|---|---|---|---|---|
| Copia DEMO (Drive) | SaaS en Render, administración | 20:33:49, 15.1 s | 20:34:35, 14.9 s | 19:48:35, 12.0 s |
| Copia local (PC) | Terminal, `--local` | EVT-000023 | EVT-000024 | EVT-000025 |
| Excel MAESTRO (Drive) | Terminal, sin bandera | EVT-000047 | EVT-000048 | EVT-000049 |

- Efecto (simulado por Claude sobre los adjuntos y visto en Render): costo unitario 'costo cupcakes'!D37 248.05 → 283.05
  en la copia DEMO y 247.00 → 282.00 en el maestro (hornada cargada de 30); 'gasto operacional' fila 9 Planilla 371,500 →
  **340,000** (ene-jun) y 382,645 → 350,200 (jul-dic), que es exactamente el sueldo del administrador; gasto operacional
  del año 16,522,417.18 → 16,138,747.18; 'TABLA CRUCE MAESTRA' G9 = 391.92; Combo 2 1,405.19 → 1,492.96.
- Por qué 343.16 era viejo: 343.16 = 217.19 (costo de 'Historico Diario') × 1.58. El costo subió a 248.05 (m005) y a 283.05
  (20 %). Con 343.16 quedaban 60.11 por unidad (21 %); con 391.92 quedan 108.87 (38 %). Con el costo de hoy, el mismo
  margen daría 447.22 (dato dado al usuario; eligió 391.92).
- Foto previa de la copia local: `datos_prueba\maestro_prueba_ANTES_cupcake20.xlsx`. Lo ya vendido y horneado no cambia.
- Por confirmar con los adjuntos del chat nuevo: los eventos del maestro y de la copia DEMO y los valores de arriba.

### 7) Mediciones (dato para el reporte)
- **PC, SaaS:** precio de referencia 8.6, 8.4, 9.0 y 9.1 s; proporción 9.1 y 9.7 s; planilla 8.5 y 8.7 s; gasto operacional
  8.7 y 8.9 s; activos 8.7 s. Prueba completa del SaaS: cerca de 8 minutos.
- **Render Starter, SaaS:** precio de referencia 12.0, 13.8, 11.6 y 11.6 s; proporción 15.1 y 14.9 s (lee una segunda vez
  para el aviso del costo). Despliegues "live": 621dfae ~7:46 p. m.; ad390ac 8:31 p. m.
- **Hornada mínima que cubre el costo** (copia DEMO): Baguette 36 unidades (costos del día 22,453.33; materiales 205.70 por
  unidad; precio 842.69). CupCake 19 con los datos finales (20 % y 391.92).
- Tiempo de trabajo (aproximado): lectura y adjuntos ~25 min; paquete 1 ~1 h 30 min (estimado 1 h 30); paquete 2 con
  proporciones y permisos ~2 h en paralelo con las pruebas del usuario (estimado 3 a 3.5 h); CupCake ~20 min.

### 8) Estado de los Excel al cierre
- **Excel maestro (Drive):** cambió solo en 'costo cupcakes' B24, B23 y G27 (EVT-000047 a EVT-000049). Lo demás igual que
  en la V23. Pedidos de prueba WA-121 y WA-122 sin cambio.
- **Copia DEMO (Drive):** igual al maestro en 'costo cupcakes' (0.20, 0.20 y 391.92). Además: tres eventos de precio de
  Baguette (queda 842.69) y lo anotado en la V21 a la V23 (fila de historia 24, existencia mínima de Baguette en 5, leche a
  1,102.50, clima fijo).
- **Copia local (PC):** EVT-000023 a EVT-000025; foto previa `maestro_prueba_ANTES_cupcake20.xlsx`.

### 9) Reglas de trabajo nuevas (van al prompt)
- Al pegar un bloque, la Terminal ejecuta todos los renglones menos el último, que espera Enter. Antes de pegar la línea
  debe estar vacía (Esc); después de pegar, Enter una vez.
- La lista de GitHub Actions no muestra los commits: se comprueban con `git --no-pager log --oneline -2` y
  `git status --short`.
- Un valor de configuración (precio de referencia, proporción, sueldo, gasto) se cambia con su función, que deja el
  evento; la migración es para estructura, fórmulas y textos.
- Un valor "raro" en la copia DEMO puede ser una decisión del usuario: preguntar antes de proponer devolverlo.
- Las pantallas de costeo (hojas de costo, planilla, gasto operacional) son solo para administración.

### 10) Libro de mapas V26 (LIBRO_MAPA_DE_CELDAS_PFV26.xlsx, **316194** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- Pestaña nueva **"Cambios V26"**. Bloques "ACTUALIZACIÓN V26" en los mapas 'costo pan baguette' (A79), 'costo cupcakes'
  (A74), 'planilla' (A19), 'MAPA gasto operacional' (A59) y 'valor activos' (A42). Notas "|| V26" en las Reglas #9 y #39.
- Decisiones Pendientes: #57, #78, #96 (RESUELTO) y #97 (aclarado) con "|| V26"; nuevos **#99** (hornada mínima, por
  diseño), **#100** (proporciones), **#101** (CupCake), **#102** (permisos), **#103** (informativo), **#104** (pendientes
  menores) y **#105** (lección).
- Verificado contra V25 celda por celda: 7,843 iguales, 6 con texto agregado "|| V26", 42 nuevas + pestaña nueva
  (62 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos. 59 pestañas.

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_actual.zip` (1,095,371; marca 1274aa1: hay que rehacerlo), `paquete_grupo4_costo.zip` (41,626),
  `paquete_grupo4_parte2_v4.zip` (54,866; las versiones de 54,336, 54,361 y 54,761 no se usaron), `ANTES_grupo4_costo.zip` y
  `ANTES_grupo4_parte2.zip` (carpetas saas_demo y microservicios antes de cada paquete).
- ROBOT PANADERO: `datos_prueba\maestro_prueba.xlsx` con CupCake en 20 % y 391.92; foto previa
  `maestro_prueba_ANTES_cupcake20.xlsx`. Sigue `saas_demo.zip` sin seguimiento.
- Drive: el maestro y la copia DEMO cambiaron como dice la sección 8. `.env` y Render: sin cambios de configuración (dos
  despliegues por los dos push).

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V24 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V24.md`) y rehacer `repo_actual.zip`.
2. Comprobar con los adjuntos: en el maestro EVT-000047 a EVT-000049 y 'costo cupcakes' B23 = B24 = 0.2 y G27 = 391.92; en
   la copia DEMO los eventos CONFIG_PRECIO_REFERENCIA y CONFIG_PROPORCION_COSTO y los mismos tres valores.
3. **#57, segundo día:** captura de GitHub Actions del 6-oct (corridas #11 y #12). Seguir viendo Ma01 (#90) y L01 (#91).
4. **Grupo 5 del SaaS:** #66 (caja chica), #67 (salidas de efectivo), #68 (flujo caja proyecto) y #69 (punto de
   equilibrio), con el plan y las celdas como evidencia; y el Registro de eventos (lectura). Con el Grupo 5, revisar el
   #104 (c): agua y electricidad en 'gasto operacional' y en las hojas de costo.
5. Decisiones del usuario: #104 (a) texto de 'DiccionarioDatos' (migración de 30 min), #104 (d) costo unitario visible para
   despacho en Hornadas, existencia mínima de Baguette en la copia DEMO, #83 y #86 (b) a (d).
6. Después: Sesión 3 (Power BI básico y pendientes menores #51, #52, #4, #17, #19, #22, #21) y Sesión 4 (entrega final),
   igual que en la V19. Durante la revisión de los tutores no se hace push (#77).
7. Del usuario: renovar los dos respaldos de RESPALDO_FUENTES antes del **13-oct-2026**.
8. Tiempo programado por el usuario: 5 días para la entrega (3 para lo técnico y 2 para la parte escrita). El 5-oct fue el
   primer día técnico (tres chats).

### Tips aprendidos hoy
- "Costo fijo del día entre unidades": mientras más se hornea, más barato sale cada pan. Por eso existe una hornada
  mínima que cubre el costo, y por eso el precio de venta es el de referencia y no costo × margen.
- Un precio de referencia es una decisión que hay que revisar cuando el costo cambia: el 343.16 era correcto con el costo
  viejo y dejó de serlo sin que nadie lo tocara.
- Dos hojas pueden depender del mismo número: la proporción del panadero mueve el costo del producto y, al revés, el gasto
  administrativo. Antes de dejar cambiar un dato hay que buscar quién más lo lee.
- Pegar varios comandos a la vez es cómodo, pero el último siempre espera Enter.
- Esconder un botón no es un permiso: el permiso se revisa en el servidor (por eso despacho recibe 403 aunque escriba la
  dirección a mano).

---

# CONTENIDO V23 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V23

*Última actualización: 5-oct-2026 (~5:15 p. m. CR). Reemplaza a la V22. La V22 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V22", que a su vez trae la V21 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V23".
Subir este archivo y el libro de mapas V25 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.** (Las claves del SaaS están solo en el `.env`, en Render y en
el libro de mapas, Regla #15.)

*Nota de fechas: el chat fue el 5-oct-2026, de las 2:15 p. m. a ~5:15 p. m. (hora CR); es el SEGUNDO chat del día (el de la
mañana es la V22). Las horas de las pruebas salen de la Terminal del usuario, de 'RegistroEventos' y de los Logs de Render.*

---

## ACTUALIZACIÓN V23 (5-oct-2026, tarde, ~3 h) — Migración m015: la margarina Ma01 se vende por caja (#36 resuelto) y 'DiccionarioDatos' al día con C16 (#86 a); SaaS Grupo 3b paquete 3: composición de los combos desde el SaaS, construida y probada en Render (#85). CIERRA EL GRUPO 3 DEL SaaS

### Estado general
- **Migración m015 CERRADA** (decisión e idea del usuario): la margarina Ma01 se vende **por caja**. El precio de venta
  'TABLA CRUCE MAESTRA'!G12 pasa de `costo / 4 × 1.05` a `costo × 1.05` (262.50 → **1,050.00**). Aplicada con doble revisión
  en la copia local, en el **Excel maestro** y en la **copia DEMO**. En GitHub (**75858fe**). **#36 RESUELTO.**
- **#86 (a) HECHO** en la misma m015: 'DiccionarioDatos' ya dice que 'NODO INTEGRAL'!C16 la escribe el robot.
- **Grupo 3b, paquete 3 (composición de los combos, #85) CERRADO:** función nueva `combos_composicion.py`, tabla de lectura
  "Qué lleva cada combo" y segundo formulario por combo en la pantalla "Tabla combos dinámicos". Probado en la PC (TODO OK),
  en GitHub (**af82dd6**) y en Render. **Con esto el Grupo 3 del SaaS queda completo.**
- **#57:** sin dato nuevo (mismo día 5-oct). El segundo día de los dos horarios se revisa el 6-oct.
- El chat se cortó aquí por decisión del usuario (recomendación de Claude: un grupo del SaaS por chat). El Grupo 4 (costeo)
  empieza en el chat nuevo.
- Libro de mapas **V25**: pestaña "Cambios V25"; #36, #57, #65, #85 y #86 con "|| V25"; nuevos **#95 a #98**.

### Commits de este chat (todos con push)
- **116bc37** — bitácora V22 (hecho al cierre del chat de la mañana; es la marca del `repo_actual.zip` adjunto).
- **75858fe** — "Migracion m015: la margarina Ma01 se vende por caja (#36) y DiccionarioDatos al dia con C16 (#86 a)":
  1 file changed, 258 insertions. Push `116bc37..75858fe`.
- **af82dd6** — "SaaS Grupo 3b paquete 3: composicion de los combos desde el SaaS (#85)": 7 files changed, 1055 insertions,
  4 deletions; 2 archivos nuevos. Push `75858fe..af82dd6`.
- **Último commit de código: af82dd6.** Pendiente al cierre: commit de esta bitácora V23 (solo este archivo).

### 1) Revisión de los adjuntos al inicio
- `repo_actual.zip`: marca **116bc37**, 985,190 bytes, 128 archivos, sin `.env` ni credenciales. Bitácora V22 230,765 bytes;
  mapa V24 300,775.
- Excel maestro: **253,165 bytes**, 25 hojas, 0 errores, último evento **EVT-000045 MIGRACION m014** (5-oct 12:52:36).
- Copia DEMO: **255,697 bytes**, 25 hojas, 0 errores, último evento **EVT-000056 MIGRACION m014** (5-oct 12:54:12).
- **m014 comprobada en los dos archivos** (pendiente 3 de la V22): `TablaComposicionCombos` = A15:H23, F16 = `H16+D16*(E16-G16)`,
  G y H en 0; Combo 1 con 5 vendidos (15 Baguette y 5 leches).
- Captura de GitHub Actions: la misma de la mañana (corridas #9 a las 00:43 y #10 a las 06:35). No hay segundo día todavía.
- **Observación (#97):** en la copia DEMO hay un evento que la V22 no menciona: **EVT-000055** (5-oct 11:54:26),
  existencia mínima de Baguette 'Control de Inventario'!G5 de 10 a 5, desde el SaaS (administración). No se devolvió.

### 2) La margarina Ma01 (#36) — cómo se decidió
- Al confirmar el plan del paquete 3, el usuario pidió explicar la margarina antes de construir. Evidencia presentada (igual en
  el maestro y en la copia DEMO): nombre "Margarina Numar caja de 4 Uds - 500 g"; costo 'TABLA CRUCE MAESTRA'!D12 = 1,000
  (la CAJA, lo que lee el robot en la API de catálogo de MaxiPalí); precio G12 = D12 / 4 × 1.05 = 262.50 (UNA barra);
  'Control de Inventario' fila 9: compras 10 − ventas 3 = existencia 7, todo contado como si fuera la misma unidad; el bot
  mostraba el nombre de la caja con el precio de la barra.
- Opciones de Claude: (A) dejar los 6 productos de reventa en la lista de los combos y el #36 como pendiente, 0 min,
  recomendada; (B) sacar la margarina de la lista, 5 min; (C) pasar todo a barras, como el café, 1.5 a 2 h.
- **Idea del usuario, que resultó la mejor:** "¿y si en vez de vender las unidades separadas se vende por caja? Creo sería
  más rápida la modificación y no afecta el costo". Comprobado con las celdas: todo lo demás YA estaba en cajas (nombre,
  costo, compras, inventario); **ninguna otra fórmula del libro lee G12**; la única venta (7-mar, 3 a 950) ya fue a precio de
  caja y está guardada como número (Regla #40); el código no tiene ninguna lógica "entre 4".
- La panadería ya no vende barras sueltas; si algún día quisiera, sería un producto nuevo con su código (como el café, m006).

### 3) Decisiones del usuario (en orden)
1. Confirmó el plan del paquete 3 con los **6 productos 'Tercero'** de la TCM en la lista del segundo componente.
2. **La margarina se vende por caja** → migración **m015** (40 min), con el **#86 (a)** en la misma migración.
3. La m015 va **antes** del paquete 3.
4. Prueba en Render con la escritura 2 (leche por margarina): aceptó que la copia DEMO quede con la fila de historia.
5. **Cerrar el chat aquí** con los tres documentos; el Grupo 4 va en el chat nuevo.
- Decisiones tomadas después del cierre de la V22 (no estaban en esa bitácora): **C8 "Acción Meta Ads" sigue MANUAL y es una
  prevista** (todavía no hay campaña; la decisión de subir la publicidad depende del movimiento del negocio y sirve igual
  para otra plataforma, p. ej. YouTube); se deja en "Sin definir" hasta que exista una campaña; no se cambia el nombre de la
  celda; se ampliará la nota de la pantalla NODO INTEGRAL. **Falta decidir** si la lista lleva también "Bajar" (#96).

### 4) Migración m015 (`migraciones/m015_margarina_por_caja.py`, **13993** bytes, 258 renglones)
- **Qué hace:** A) 'TABLA CRUCE MAESTRA', fila de Ma01, columna 'Precio': de
  `=TablaProductos[[#This Row],[Costo Unitario (Robot)]]/4*(1.05)` a `=TablaProductos[[#This Row],[Costo Unitario (Robot)]]*(1.05)`
  (la misma fórmula de los otros cinco productos de reventa). B) 'DiccionarioDatos' N234 (COST.PROD.COSTO_MARGARINA): nota
  agregada. C) 'DiccionarioDatos' G257, H257 y N257 (NODO.ACTIVACION_PROMOCION): C16 no tiene fórmula, la escribe el robot
  (ON = algún combo ACTIVO o descuento general de día de pago / feriado). D) Evento MIGRACION. Nada se borra: lo nuevo va
  después de "|| m015".
- Ubica la fila por el CÓDIGO (Ma01) y las del diccionario por su ID, no por número fijo. Se detiene sin escribir si la
  fórmula no es la esperada, si el costo no es un número mayor que 0 o si la migración está aplicada a medias. Banderas
  `--local`, sin bandera (maestro) y `--demo`, igual que la m014.
- **Resultados (5-oct-2026):**

| Archivo | Hora | Resultado |
|---|---|---|
| Copia local | 2:46 p. m. | LISTO (A 1, B 1, C 3 celdas); evento EVT-000022; 2a corrida "Ya estaba aplicada"; comparar_libros: 16 celdas distintas (1 TCM, 4 DiccionarioDatos, 11 del evento), **0 valores calculados**, 0 errores nuevos |
| Excel maestro (Drive) | ~2:56 p. m. | LISTO; 2a corrida "Ya estaba aplicada"; probar_conexion TODO BIEN (**versión 222** de Drive; Ma01 Precio 1050) |
| Copia DEMO (Drive) | ~2:58 p. m. | LISTO; 2a corrida "Ya estaba aplicada" |
- En `comparar_libros` el precio de la margarina sale como fórmula cambiada (no como "valor calculado"): por eso el renglón
  dice 0. En el entorno de Claude, sobre los dos archivos de Drive adjuntos: mismos resultados (eventos EVT-000046 en el
  maestro y EVT-000057 en la copia DEMO: son los números esperados en Drive, por confirmar con los adjuntos del chat nuevo);
  'Control de Inventario' fila 9 sin cambio (10, 3, 7); el catálogo (bot, SaaS y despacho) ya lee 1,050.
- Foto previa en la PC: `datos_prueba\maestro_prueba_ANTES_m015.xlsx` (Git no la sigue).

### 5) Lo construido — paquete 3 (`paquete_grupo3b_composicion.zip`, **40245** bytes, 7 archivos)
**Nuevos (2):** `microservicios/grupo3_integraciones/combos_composicion.py` (**26179**; `leer_composicion()`,
`normalizar_cantidad()`, `formula_precio()`, `leer_formula_precio()` y
`actualizar_composicion(codigo, producto_1, cantidad_1, producto_2, cantidad_2, origen)`) y `probar_combos_composicion.py`
(**23277**).

**Modificados (5), solo con agregados:** `saas_demo/grupo3.py` (**17503**, antes 14178; la pantalla recibe la composición y
ruta nueva `/combos/composicion/guardar`), `saas_demo/plantillas/combos.html` (**14415**, antes 8174),
`saas_demo/probar_saas_demo.py` (**60569**, antes 52977; sección O), `saas_demo/__init__.py` (1088) y
`grupo3_integraciones/__init__.py` (1528). No se tocaron `combos_config.py`, `catalogo.py` ni `app.py`.

Cómo trabaja `actualizar_composicion` (un guardado = UNA escritura y UN evento **COMPOSICION_COMBO**):
- Recibe el combo por CÓDIGO y los productos por código o nombre de la TCM. El primero debe ser Tipo 'Propio' y el segundo
  'Tercero'; cantidades enteras de 1 a 10. Si no, rechaza y no escribe.
- **Antes de escribir comprueba que la fórmula del precio (E) y la tabla de composición digan lo mismo** (mismas cantidades
  y mismas filas de la TCM). Si no coinciden, se detiene (los mapas y el Excel mandan; no se improvisa).
- **Cambia solo la cantidad:** G := combos vendidos (E), H := unidades rebajadas (F), y luego D := cantidad nueva. F no se mueve.
- **Cambia el producto:** si había unidades rebajadas, van a una **fila de historia** al final de la tabla (mismas fórmulas en
  A, B, E y F; C = producto anterior; D = 0; H = lo rebajado); si ya existe una para ese combo y producto, se acumula en
  ella; si no había nada rebajado, no se crea. La fila vigente pasa al producto nuevo: C enlace a la TCM, D cantidad nueva,
  G := E, H := 0. La tabla crece sola (`agregar_fila`, ya probada) y 'Control de Inventario'!I la suma por su nombre.
- En el mismo guardado: la fórmula de E `=(n1*'TABLA CRUCE MAESTRA'!Gf1)*(1-IF($L>1,$L/100,$L))+n2*'TABLA CRUCE MAESTRA'!Gf2`
  (descuento solo sobre el componente propio, Regla #28); el detalle D con los nombres de la TCM (Regla #26); y el mensaje K,
  si traía el detalle viejo tal cual (si pasara de 150 caracteres o no lo traía, no se toca y la pantalla lo avisa).
- Sin cambios: SIN_CAMBIO, no sube nada. No revisa C16: la composición no activa ni desactiva combos.
- Uso manual: `py -m microservicios.grupo3_integraciones.combos_composicion --local --ver`.

Pantalla "Tabla combos dinámicos": tabla nueva **"Qué lleva cada combo (composición)"** para los dos usuarios (componente,
producto, cantidad, combos vendidos, unidades rebajadas, fila; las filas de historia salen como "Historia (ya no lo lleva)")
y, solo para administración, un formulario **"Combo N: qué lleva"** por combo con el botón "Guardar composición de Combo N".

**Límite conocido (#92), sin cambio:** se congela por contador de combos, no por fecha; anular después de un cambio una
venta anterior al cambio devuelve la composición nueva.

### 6) Pruebas
- **Entorno de Claude** (copia DEMO adjunta, ya con la m015): `probar_combos_composicion` TODO OK; `probar_saas_demo`
  **162 revisiones** (antes 149), **24 escrituras**, TODO OK. Las pruebas atraparon dos errores de Claude durante la
  construcción, corregidos antes de entregar: la lectura de la fórmula del precio fallaba porque el nombre de la hoja tiene
  espacios, y un choque de nombres en la sección O de la prueba.
- **PC del usuario (~4:15 a 4:28 p. m.):** los 7 archivos con su tamaño; `probar_combos_composicion` **RESULTADO: TODO OK**
  (en la copia local el Combo 1 no tenía ventas: la prueba vendió 2, congeló 2 combos y 6 Baguette; precio 3,377.76 →
  4,136.18; fila de historia 24, tabla A15:H23 → A15:H24; la historia se acumuló 3 + 1 = 4; devolvió todo);
  `probar_saas_demo` **RESULTADO: TODO OK**, secciones A a O; `datos_prueba\maestro_prueba.xlsx` intacta.
- **Render (4:52 a 5:01 p. m., administración, copia DEMO, Combo 1):**

| Escritura | Hora (Logs) | Tiempo | Resultado en pantalla |
|---|---|---|---|
| Cantidad 3 → 4 Baguette | 16:52:40 | 12.0 s | Baguette 4, 5 vendidos, 15 rebajadas; leche igual (1, 5, 5) |
| Leche → Margarina (cantidad 1) | 16:57:44 | 11.4 s | detalle "4 Baguette + 1 Margarina Numar caja de 4 Uds - 500 g"; margarina 0 unidades; fila de historia 24 con 5 leches |
| De vuelta a 3 Baguette + 1 Leche | 17:01:09 | 11.8 s | formulario en 3 y 1; "ya rebajado: 15 Baguette y 0 Leche" (las 5 siguen en la fila de historia) |
- El inventario no se movió: 15 Baguette y 5 leches. **En las capturas no salió el aviso verde** (sale arriba de la página):
  los tiempos se leyeron en los Logs de `saas-panaderia`.
- Los Logs muestran además **dos guardados de CONFIGURACIÓN del Combo 1** (16:47:14 y 16:52:01; 15.2 y 15.3 s) que no eran
  parte del plan: en las capturas la configuración se ve como al inicio (TEMPERATURA <= 30, prioridad 1, 10 %). Por confirmar
  con los eventos CONFIG_COMBO de la copia DEMO (#97).

### 7) Mediciones (dato para el reporte)
- **PC, SaaS:** guardar la composición de un combo 8.1 s (tres veces). Prueba del SaaS completa: 24 escrituras, ~5 min 50 s.
- **Render Starter, SaaS:** composición 12.0, 11.4 y 11.8 s; configuración de un combo 15.2 y 15.3 s (incluye la revisión de
  C16; es el dato que faltó medir en la V22). Despliegue del af82dd6 "live" a las 4:40 p. m.
- **m015:** tres archivos migrados entre las 2:46 y ~2:58 p. m.; maestro en la versión 222 de Drive.
- Tiempo de trabajo (aproximado): lectura y revisión de adjuntos ~20 min; margarina y m015 ~45 min (estimado 40);
  paquete 3 ~1 h 45 min con instalación, pruebas y commit (estimado 2 h); Render ~20 min.

### 8) Estado de los Excel al cierre
- **Excel maestro (Drive):** solo cambió por la m015 (G12 y 4 textos de 'DiccionarioDatos'; evento esperado EVT-000046).
  `TablaComposicionCombos` sigue A15:H23 con G y H en 0. Pedidos de prueba WA-121 y WA-122 sin cambio.
- **Copia DEMO (Drive):** m015 (evento esperado EVT-000057), dos CONFIG_COMBO y tres COMPOSICION_COMBO. Queda:
  `TablaComposicionCombos` **A15:H24**; fila 16 Baguette D = 3, G = 5, H = 15; fila 17 Leche D = 1, G = 5, H = 0;
  **fila 24** historia de la Leche (D = 0, H = 5); E5 con la fórmula nueva (termina en `+1*'TABLA CRUCE MAESTRA'!G10`) y el
  mismo precio 3,377.76; D5 y K5 con el texto de siempre. Además: lo anotado en la V21 y la V22, la existencia mínima de
  Baguette en 5 (EVT-000055), la leche a 1,102.50 (en el maestro 1,029) y el clima fijo.
- **Copia local (PC):** tiene la m015 (EVT-000022). Las pruebas trabajaron sobre copias temporales.

### 9) Reglas de trabajo nuevas (van al prompt)
- Los pasos para la Terminal van en UN bloque por mensaje y al final; si Claude va a seguir trabajando, lo dice después de
  los pasos y no agrega otros pasos en el mismo mensaje (el usuario repitió pasos ya cerrados; no pasó nada).
- Si `Compress-Archive` dice que el respaldo ya existe, no se usa `-Force`: se comprueba con `tar -tf ... | Select-String`
  que el respaldo es el de ANTES.
- El aviso verde sale ARRIBA de la página: la captura se toma sin bajar. Si no salió, el tiempo se lee en los Logs de
  `saas-panaderia` buscando la palabra de la escritura ("composicion", "combo").
- Antes de proponer una corrección grande, revisar si basta con cambiar el lado contrario (lección de la margarina).

### 10) Libro de mapas V25 (LIBRO_MAPA_DE_CELDAS_PFV25.xlsx, **307536** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- Pestaña nueva **"Cambios V25"**. Bloques "ACTUALIZACIÓN V25" en los mapas 'Tabla combos dinamicos' (A106) y
  'Mapa TABLA CRUCE MAESTRA' (A68). Notas "|| V25" en las Reglas #32 y #40 y en el mapa 'NODO INTEGRAL' (C8 y C16).
- Decisiones Pendientes: #36 (RESUELTO), #57, #65 (completo), #85 (construido) y #86 con "|| V25"; nuevos **#95**
  (informativo: evidencia y tiempos), **#96** (pendiente menor: nota de C8 y opción "Bajar"), **#97** (observación: copia
  DEMO) y **#98** (lección).
- Verificado contra V24 celda por celda: 7,764 iguales, 11 con texto agregado "|| V25", 18 nuevas + pestaña nueva
  (56 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos. 58 pestañas.

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_actual.zip` (985,190; marca 116bc37: hay que rehacerlo), `m015_margarina_por_caja.py` (13,993),
  `paquete_grupo3b_composicion.zip` (40,245; 7 archivos) y `ANTES_grupo3b_composicion.zip` (carpetas saas_demo y
  microservicios antes del paquete 3; comprobado con `tar -tf` que no trae `combos_composicion.py`).
- ROBOT PANADERO: `datos_prueba\maestro_prueba.xlsx` ya tiene la m015; foto previa `maestro_prueba_ANTES_m015.xlsx`. Sigue
  `saas_demo.zip` sin seguimiento (se puede mover a download).
- Drive: el maestro cambió por la m015; la copia DEMO, por la m015 y por la prueba en Render (sección 8). `.env` y Render: sin
  cambios de configuración (dos despliegues por los dos push).

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V23 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V23.md`).
2. Comprobar con los adjuntos: evento MIGRACION m015 y G12 = 1,050 en el maestro y en la copia DEMO; en la copia DEMO, los
   tres COMPOSICION_COMBO, la tabla A15:H24 y qué cambiaron los dos CONFIG_COMBO de las 4:47 y 4:52 p. m. (#97).
3. **#57, segundo día:** captura de GitHub Actions del 6-oct. Seguir viendo Ma01 (#90) y, opcional, el renglón de L01 en el
   paso "2) Precios de proveedores" (#91).
4. **Grupo 4 del SaaS (costeo):** plan con las celdas como evidencia y construcción.
5. Menor (#96): nota de C8 en la pantalla NODO INTEGRAL y decidir si la lista lleva "Bajar"; va con el primer paquete del
   Grupo 4. Menor (#97): devolver o no la existencia mínima de Baguette a 10 en la copia DEMO.
6. Luego Grupo 5 (#66 a #69) y el Registro de eventos (lectura). Decisiones del usuario: #83 (menor) y #86 (b) a (d).
7. Después: Sesión 3 (Power BI básico y pendientes menores #51, #52, #4, #17, #19, #22, #21) y Sesión 4 (entrega final),
   igual que en la V19. Durante la revisión de los tutores no se hace push (#77).
8. Del usuario: renovar los dos respaldos de RESPALDO_FUENTES antes del **13-oct-2026**.
9. Tiempo programado por el usuario: 5 días para la entrega (3 para lo técnico y 2 para la parte escrita), con chats de
   unas 6 horas. Este fue el primer día técnico (dos chats).

### Tips aprendidos hoy
- A veces la corrección más limpia es la más corta: si casi todo está en una unidad (cajas) y una sola celda en otra (barras),
  se cambia esa celda y no todo lo demás.
- Una fórmula puede escribirla un programa: el robot arma el texto `=(4*...)` con las cantidades nuevas y el Excel la calcula.
- "Fila de historia": cuando un dato deja de usarse pero lo ya hecho debe seguir contando, no se borra ni se cambia; se guarda
  aparte con cantidad 0 para que la suma siga dando lo mismo.
- Antes de cambiar algo que vive en tres lugares (texto, fórmula y tabla), el programa comprueba que los tres digan lo mismo.
- Cuando un comando dice "ya estaba aplicada" o "nothing to commit", repetirlo no hizo daño: a eso se le llama ser idempotente.

---

# CONTENIDO V22 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V22

*Última actualización: 5-oct-2026 (~1:10 p. m. CR). Reemplaza a la V21. La V21 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V21", que a su vez trae la V20 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V22".
Subir este archivo y el libro de mapas V24 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.** (Las claves del SaaS están solo en el `.env`, en Render y en
el libro de mapas, Regla #15.)

*Nota de fechas: el chat fue el 5-oct-2026, de las 9:30 a. m. a ~1:10 p. m. (hora CR). Las horas de las pruebas salen
de la Terminal del usuario, de GitHub Actions y de 'RegistroEventos'.*

---

## ACTUALIZACIÓN V22 (5-oct-2026, mañana, ~3 h 40 min) — SaaS Grupo 3b: configuración de combos construida y probada en Render (#65); catálogo corregido (#89); migración m014 aplicada en el maestro y en la copia DEMO (#85); #57, #58 y #84 comprobados

### Estado general
- **Grupo 3b, paquete 1 (configuración de combos) CERRADO:** función nueva `combos_config.py`, pantalla "Tabla combos
  dinámicos" (menú 3 Integraciones), probado en la PC (149 revisiones, TODO OK), en GitHub (**beed350**) y en Render.
- **#89 nuevo y resuelto:** `catalogo.py` tomaba las filas de `TablaComposicionCombos` como combos; ahora la lista de combos
  termina en la primera fila vacía.
- **Grupo 3b, paquete 2 (migración m014) CERRADO:** aplicada en la copia local, en el **Excel maestro** de Drive y en la
  **copia DEMO** de Drive, con doble revisión. En GitHub (**fb99f10**). No cambió ningún valor calculado.
- **Grupo 3b, paquete 3 (composición ajustable desde el SaaS) NO construido:** diseño confirmado por el usuario; es lo
  primero del chat nuevo (2 h).
- **#57:** con los dos horarios, las corridas "Scheduled" del 5-oct salieron a las 12:43 a. m. y a las 6:35 a. m. (26 y 18
  minutos de atraso; el 3 y el 4 de octubre fueron 5 a 6 horas). Un solo día: sigue en observación.
- **#58 VISTO por primera vez en la nube** (respaldo manual de Solera) y **#84 comprobado en el maestro** (C16 = ON).
- **#81 CERRADO por decisión del usuario:** los mínimos los administra el negocio desde el SaaS; no se cambia nada.
- Libro de mapas **V24**: pestaña "Cambios V24"; #57, #58, #65, #81, #84, #85 y #86 con "|| V24"; nuevos **#89 a #94**.

### Commits de este chat (todos con push)
- **1809a05** — bitácora V21 (hecho al cierre del chat anterior).
- **beed350** — "SaaS Grupo 3b paquete 1: configuracion de combos (#65) y catalogo que lee solo la lista de combos (#89)":
  9 files changed, 948 insertions, 3 deletions; 3 archivos nuevos. Push `1809a05..beed350`.
- **fb99f10** — "Migracion m014: columnas para congelar lo ya rebajado al cambiar la composicion de un combo (#85)":
  1 file changed, 279 insertions. Push `beed350..fb99f10`.
- **Último commit de código: fb99f10.** Pendiente al cierre: commit de esta bitácora V22 (solo este archivo).

### 1) Revisión de los adjuntos al inicio
- `repo_actual.zip`: marca **1809a05**, 880,684 bytes, 123 archivos, sin `.env` ni credenciales. Bitácora V21 209,397
  bytes; mapa V23 294,425.
- Excel maestro: **251,499 bytes**, 25 hojas, 0 errores, 44 eventos (último EVT-000044 del 5-oct 06:38).
- Copia DEMO: **253,564 bytes**, 25 hojas, 50 eventos (último EVT-000050, tarifa 550 → 500). El usuario adjuntó después por
  error la copia anterior (253,288 bytes, 47 eventos, C16 con el relleno): no se usó.
- Capturas: GitHub Actions (10 corridas) y, más tarde, NODO INTEGRAL en Render ya sin el botón (pendiente de la V21: cerrado).

### 2) #57, #58 y #84 — lo que dejó la tarea diaria del 5-oct en el maestro
| Evento | Hora (CR) | Qué dice |
|---|---|---|
| EVT-000040 | 00:44:00 | ACTIVACION_PROMOCION: 'NODO INTEGRAL'!C16 '[ON / OFF]' → 'ON', combos activos co01, co04 (**#84**) |
| EVT-000041 | 00:44:34 | PRECIOS_PROVEEDORES **PARCIAL**: cambiado L01 TCM!D10 1050 → 980; sin cambio 14; fallido Ma01; **"Desde RESPALDO manual 11: respaldo_solera.json del 03-10-2026 (2 dias)"** (**#58**) |
| EVT-000042 | 00:44:47 | COMPETENCIA OK: Musi sin promoción, fecha 05-10-2026 |
| EVT-000043 | 06:37:59 | PRECIOS_PROVEEDORES PARCIAL: sin cambio 15; fallido Ma01. Solera se leyó DIRECTO (el bloqueo es intermitente) |
| EVT-000044 | 06:38:12 | COMPETENCIA OK, sin cambios |
- **#57 (captura de GitHub Actions):** corrida **#9 Scheduled a las 00:43** (programada 12:17 a. m., 1 min 55 s) y **#10
  Scheduled a las 06:35** (programada 6:17 a. m., 2 min 40 s), las dos en verde. La #8 del 4-oct había salido a las 12:18 p. m.
- La segunda corrida no dejó evento de C16 porque el valor no cambió (así se diseñó).
- **Hallazgo #90 (observación):** Ma01 (Margarina Numar, MaxiPalí) devolvió precio 0 en las dos corridas; el robot no tocó la
  celda (sigue en 1,000) y por eso el evento dice PARCIAL: la Regla #19 funcionó. Si se repite varios días es un caso de la
  Regla #17 (reemplazo, decide el usuario).
- **Hallazgo #91 (por comprobar):** L01 (leche, MaxiPalí) bajó el costo de 1,050 a 980 → precio de venta 1,102.50 → 1,029 y
  Combo 1 3,377.76 → 3,304.26 en el MAESTRO (Regla #28). `precios_proveedores.py` tiene para L01 un segundo código de barras
  de respaldo (otra marca) y el evento no dice cuál se leyó. El usuario confirmó que la leche es la Dos Pinos, como se
  definió. Se comprueba en GitHub Actions, paso **"2) Precios de proveedores"** (octavo de la lista), renglón que empieza
  con L01: si dice "(RESPALDO)" el 980 es de la otra marca. La copia DEMO sigue con la leche a 1,102.50.

### 3) Grupo 3b — evidencia de celdas ('Tabla Combos Dinamicos'; mapa y Excel coinciden, salvo el #86)
- Filas 5 a 8: A Detonador (fórmula), B texto del criterio, C nombre (llave), D detalle, E precio (fórmula), F código,
  **G tipo de detonador, H operador, I valor límite, J prioridad, K mensaje, L descuento %** (las seis manuales).
- La composición de un combo vive en TRES lugares: D5 (texto "3 Baguette + 1 Leche..."), la fórmula de E5
  (`(3*'TABLA CRUCE MAESTRA'!G8)*(1-IF($L5>1,$L5/100,$L5)) +'TABLA CRUCE MAESTRA'!G10`, cantidades y filas escritas dentro)
  y `TablaComposicionCombos` (A15:F23, dos filas por combo: C producto como enlace a la TCM, D cantidad, E combos
  vendidos con SUMIFS de 'Ventas despacho' y 'Detalle comanda digital' sin fecha, F = D × E).
- 'Control de Inventario' columna I = SUMIF de la columna F por producto. Solo el **Combo 1** tiene ventas: 5, todas
  digitales (WA-119, WA-120, WA-121) → 15 Baguette y 5 leches. Igual en el maestro y en la copia DEMO.
- El código NO lee la composición: `catalogo.py` lee de la fila 4 hacia abajo por encabezado (nombre, código, precio,
  estado, detalle). El inventario lo rebaja el Excel con sus fórmulas.
- La hora de un pedido digital SÍ se guarda (aclaración del usuario, según el mapa): 'Interface digital' columna C (HORA).
  'Detalle comanda digital' solo trae FECHA, enlazada por ID_PEDIDO.

### 4) Decisiones del usuario (en orden)
1. Confirmó el plan base del 3b: H lista <, <=, >, >=; I de 0 a 50 con un decimal (con LLUVIA: Lluvia / Despejado); J de 1 a
   9; K hasta 150 caracteres; B la reescribe el robot; un formulario por combo, solo administración; al guardar el robot
   revisa C16.
2. Confirmó el **diseño de la m014** (sección 7) y la **forma de la composición**: dos componentes, el primero de producción
   propia (Baguette o CupCake) y el segundo de reventa, cantidad entera de 1 a 10.
3. **Anulación (#92), opción A:** si después de un cambio de composición se anula una venta hecha ANTES del cambio, el
   inventario devuelve la composición nueva y no la vieja; queda como **límite conocido** en el reporte.
4. **Mensaje K, opción A:** al cambiar la composición, el robot cambia dentro de K el detalle viejo por el nuevo si aparece
   tal cual ("debe actualizarse").
5. **Catálogo (#89), opción A:** corregirlo ya.
6. Detalles de construcción del paquete 1: el descuento se deja como está (el SaaS muestra 1 %; la celda guarda 0.01; sin
   formato de porcentaje) y el clima de la copia DEMO queda **fijo con una nota en la pantalla** (opción A). La opción B (que
   la tarea diaria actualice también NODO INTEGRAL de la copia DEMO, #86) sigue sin aprobar.
7. **#81 cerrado:** el mínimo ya vive en 'Control de Inventario' G5:G22 y la administración lo cambia desde el SaaS; "¿qué
   sentido tiene crear una tabla para tal fin?". No se cambia ningún valor; el 10 parejo queda como dato supuesto (Regla #27).
8. "Seguir con la migración m014".

### 5) Lo construido — paquete 1 (`paquete_grupo3b_config_v2.zip`, **42583** bytes, 9 archivos; reemplazó al de 42442)
**Nuevos (3):** `microservicios/grupo3_integraciones/combos_config.py` (**18014**; `leer_combos()`,
`normalizar_configuracion(...)`, `actualizar_combo(codigo, tipo, operador, valor, prioridad, mensaje, descuento, origen)`),
`probar_combos_config.py` (**18701**) y `saas_demo/plantillas/combos.html` (**8174**).

**Modificados (6):** `microservicios/grupo1_captura_operativa/catalogo.py` (**9212**; #89), `saas_demo/grupo3.py`
(**14178**; rutas `/combos` y `/combos/guardar`), `saas_demo/app.py` (**14571**; menú "Tabla combos dinámicos"),
`saas_demo/probar_saas_demo.py` (**52977**; sección N), `saas_demo/__init__.py` (1005) y
`grupo3_integraciones/__init__.py` (1169).

Cómo trabaja `actualizar_combo`: ubica el combo por su CÓDIGO y las columnas por su ENCABEZADO; escribe solo las celdas
que cambiaron de G a L; alinea el texto de la columna B ("<= 30" o "Lluvia"); deja UN evento **CONFIG_COMBO** con
anterior → nuevo; y, ya con el Excel recalculado, llama a `actualizar_activacion_promocion` (C16). Sin cambios en G a L:
SIN_CAMBIO, no sube nada. Rechaza FERIADO y DIA_PAGO (Regla #8), operadores fuera de la lista, temperatura fuera de 0 a 50,
prioridad fuera de 1 a 9, mensaje vacío, de más de 150 caracteres o que empiece con "=", y descuento fuera de 0 a 10.

Uso manual: `py -m microservicios.grupo3_integraciones.combos_config --local --ver`.

**Tres detalles de construcción (presentados al usuario):**
- La fórmula del precio lee el descuento con `IF(L>1, L/100, L)`: un 1 sería 100 %. El robot guarda 2, 5 o 10 tal cual y
  los porcentajes de 1 o menos como fracción (1 % → 0.01). La pantalla siempre muestra el porcentaje.
- La columna B se reescribe solo cuando hay un cambio real en G a L.
- La copia DEMO tiene el clima fijo (Lluvia, 28.6 °C): la pantalla lo dice. Con LLUVIA la pantalla trae dos campos para el
  valor límite (temperatura y clima) y el robot recibe el que corresponde al detonador.

### 6) #89 — el catálogo leía de más
- `Catalogo._leer_combos` recorría TODA la hoja: las filas 15 a 23 (`TablaComposicionCombos`) tienen un producto bajo la
  columna del nombre y un número bajo la del código, y entraban como "combos". Pedir "Baguette" por NOMBRE respondía
  "El combo 'Baguette' (15) no esta ACTIVO hoy".
- No se notaba porque el bot, el SaaS y las pruebas piden por CÓDIGO (P01, Pa01).
- Corrección (9 renglones): la lista termina en la primera fila sin nombre ni código. Revisado en `probar_combos_config`
  (4 combos; "Baguette" por nombre = P01).

### 7) Migración m014 (`migraciones/m014_congelar_combos.py`, **14547** bytes)
- **Qué hace:** `TablaComposicionCombos` pasa de A15:F23 a **A15:H23** con dos columnas nuevas en 0: **G "Combos ya
  congelados"** y **H "Unidades ya rebajadas"**. La columna F pasa de `D×E` a **`H + D×(E−G)`** (mismo resultado con
  ceros). Dos filas nuevas en 'DiccionarioDatos' (309 y 310: MKT.COMBO.COMP.CONGELADOS y
  MKT.COMBO.COMP.UNIDADES_CONGELADAS) y un evento MIGRACION.
- **Banderas:** `--local` (copia de datos_prueba), sin bandera (maestro de Drive) y **`--demo`** (copia DEMO de Drive: usa
  la configuración del SaaS, no arranca si SAAS_EXCEL_FILE_ID es el maestro y exige la palabra DEMO en el nombre del archivo;
  sin `--demo` se detiene si el archivo se llama DEMO).
- **Resultados (5-oct-2026):**

| Archivo | Hora | Resultado |
|---|---|---|
| Copia local | 12:44 p. m. | LISTO (A 18, B 8, C 28 celdas); 2a corrida "Ya estaba aplicada"; comparar_libros: tabla A15:F23 → A15:H23, 26 + 28 celdas, evento EVT-000021, 65 celdas distintas, 0 errores nuevos |
| Excel maestro (Drive) | ~12:52 p. m. | LISTO; 2a corrida "Ya estaba aplicada"; probar_conexion TODO BIEN (versión 219 de Drive) |
| Copia DEMO (Drive) | ~12:54 p. m. | LISTO; 2a corrida "Ya estaba aplicada" |
- En el entorno de Claude, sobre los dos archivos de Drive descargados ese día: F16:F23 y 'Control de Inventario' idénticos
  antes y después; valores calculados que cambian: 0 en el maestro y 5 en la copia DEMO, **los cinco de la fila del día de
  'Flujo caja diario'** (fórmula con HOY, aparece al recalcular en un día nuevo; no es de la migración). En la copia local
  del usuario salieron 30, que coincide con ese patrón (5 por día).
- Foto previa en la PC: `datos_prueba\maestro_prueba_ANTES_m014.xlsx` (Git no la sigue).

### 8) Pruebas
- **Entorno de Claude:** `probar_combos_config` y `probar_saas_demo` TODO OK sobre la copia DEMO sin migrar y migrada, y
  sobre una copia con C16 todavía en "[ON / OFF]". También `probar_activacion_promocion` y `probar_ventas_despacho`.
- **PC del usuario (11:00 a 11:07 a. m.):** las dos pruebas con **RESULTADO: TODO OK (0 lineas ERROR)**;
  `datos_prueba\maestro_prueba.xlsx` intacta. Después de la m014 (12:46 p. m.), `probar_combos_config` otra vez TODO OK.
- **Render (11:20 a 11:35 a. m., administración, copia DEMO):** pantalla con "Hoy en NODO INTEGRAL: Lluvia, 28.6 °C", la
  nota del clima fijo, Oferta hoy ON (co01, co04) y la promoción de Musi; Combo 1 descuento 10 → 5 → 10 (precio otra vez
  3,377.76); Combo 2 LLUVIA = Lluvia → ACTIVO ("Clima = Lluvia") y de vuelta a TEMPERATURA >= 30, INACTIVO.
  **En las capturas no salió el aviso verde:** el tiempo de esas cuatro escrituras en Render quedó sin medir (está en los
  Logs de saas-panaderia, renglones "listo en N s").

### 9) Mediciones (dato para el reporte)
- **GitHub Actions con dos horarios (5-oct):** 26 y 18 minutos de atraso; duración 1 min 55 s y 2 min 40 s.
- **PC, SaaS:** guardar la configuración de un combo 9.4 y 9.3 s (incluye la revisión de C16, que no escribió).
- **Prueba del SaaS completa:** 149 revisiones (antes 136), 21 escrituras, cerca de 4.5 minutos.
- Tiempo de trabajo (aproximado): lectura y revisión de adjuntos ~25 min; evidencia y diseño ~30 min; paquete 1 ~1 h 30 min
  (estimado 1.5 h); m014 ~1 h (estimado 1.5 h).

### 10) Paquete 3 — lo que falta construir (diseño confirmado; estimado 2 h)
Función nueva del robot (propuesta: `microservicios/grupo3_integraciones/combos_composicion.py`, con su prueba) y un segundo
formulario por combo en la pantalla "Tabla combos dinámicos" (solo administración):
- **Entrada:** código del combo; componente 1 = producto 'Propio' de la TCM y cantidad; componente 2 = producto 'Tercero' y
  cantidad; cantidades enteras de 1 a 10. Los nombres salen de la TCM (Regla #26).
- **En UN solo guardado:**
  1. Tabla de composición, fila vigente de cada componente (las de cantidad mayor que 0). Si cambia solo la CANTIDAD:
     G := E de ese momento (combos vendidos) y H := F de ese momento (unidades rebajadas), y luego D := cantidad nueva.
     Si cambia el PRODUCTO: se agrega una **fila de historia** al final de la tabla (o se acumula en la que ya exista para
     ese combo y producto) con el producto anterior, D = 0 y H = lo ya rebajado; la fila vigente pasa al producto nuevo
     (C como enlace a la TCM), D := cantidad nueva, G := E y H := 0. Al agregar filas se extiende el rango de la tabla.
  2. La fórmula de E de ese combo se reescribe con las cantidades y las filas nuevas de la TCM (columna G = Precio); el
     descuento L sigue aplicando solo al componente propio (Regla #28).
  3. El texto D ("4 Baguette + 1 Gaseosa Coca Cola Regular - 600 ml").
  4. El mensaje K: se cambia el detalle viejo por el nuevo si aparece tal cual.
  5. Un evento con anterior → nuevo y lo congelado.
- **Ejemplo con los datos de hoy (Combo 1, 5 vendidos):** 3 → 4 Baguette: G16 = 5, H16 = 15, D16 = 4; F16 sigue en 15 y la
  siguiente venta suma 4. Leche → Coca Cola: fila de historia con Leche, D = 0, H = 5; la fila 17 pasa a Coca Cola en 0.
- **Límite conocido (#92):** el diseño cuenta combos, no fechas; anular después de un cambio una venta anterior al cambio
  devuelve la composición nueva.
- **La prueba debe revisar:** una venta posterior al cambio rebaja con la composición nueva; lo anterior no se mueve; el
  precio de las ventas ya hechas no cambia (Regla #40); segunda vez SIN_CAMBIO; devolver todo.

### 11) Reglas de trabajo nuevas (van al prompt)
- Al pedir capturas de una prueba en Render, pedir SIEMPRE la del aviso verde; si no salió, el tiempo se lee en los Logs
  del servicio ("listo en N s") sin repetir la escritura.
- Si el usuario adjunta dos copias con el mismo nombre, se comparan por el último evento y C16 antes de usar una.
- Antes de proponer algo "nuevo" para un dato, comprobar si ya existe en el Excel y en el SaaS (lección del #81) y no llamar
  "tabla" a una lista de valores.
- En `comparar_libros`, los "valores calculados que cambiaron" pueden ser la fila del día de 'Flujo caja diario' (5 por día).
- Las migraciones tienen la bandera `--demo` para la copia de demostración (desde la m014).

### 12) Libro de mapas V24 (LIBRO_MAPA_DE_CELDAS_PFV24.xlsx, **300775** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- Pestaña nueva **"Cambios V24"**. Bloque "ACTUALIZACIÓN V24" en el mapa 'Tabla combos dinamicos'.
- Decisiones Pendientes: #57, #58, #65, #81, #84, #85 y #86 con "|| V24"; nuevos **#89** (resuelto: catálogo), **#90**
  (observación: Ma01 con precio 0), **#91** (por comprobar: código de barras de L01), **#92** (límite conocido: anulación
  tras un cambio de composición), **#93** (informativo: evidencia y tiempos) y **#94** (lección).
- Verificado contra V23 celda por celda: 7,681 iguales, 7 con texto agregado "|| V24", 25 nuevas + pestaña nueva
  (62 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos. 57 pestañas.

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_actual.zip` (880,684; marca 1809a05), `ANTES_grupo3b_config.zip` (carpetas saas_demo y
  microservicios antes del paquete 1), `paquete_grupo3b_config_v2.zip` (42,583; 9 archivos), `m014_congelar_combos.py`
  (14,547). `paquete_grupo3b_config.zip` (42,442) quedó reemplazado y no se instaló. Hay varias copias DEMO viejas con
  "(1)", "(2)" y "(3)" en el nombre: la vigente es siempre la recién descargada de Drive.
- ROBOT PANADERO: `datos_prueba\maestro_prueba.xlsx` ya tiene la m014; foto previa `maestro_prueba_ANTES_m014.xlsx`. Sigue
  `saas_demo.zip` sin seguimiento (se puede mover a download).
- Drive: el **maestro** cambió por la tarea diaria (C16 = ON, leche a 980) y por la m014; la **copia DEMO**, por cuatro
  eventos CONFIG_COMBO de la prueba en Render (la configuración se devolvió; el texto de la columna B de los combos 1 y 2
  quedó "<= 30" y ">= 30") y por la m014. `.env` y Render: sin cambios.

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V22 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V22.md`).
2. **Paquete 3 del 3b:** composición ajustable (sección 10), con su prueba, commit y Render (2 h). Con eso cierra el Grupo 3.
3. Comprobar con los adjuntos: en el maestro y en la copia DEMO, tabla A15:H23 y evento MIGRACION m014; segundo día del #57.
4. Opcionales: captura de los Logs de saas-panaderia con los tiempos de combos; renglón de L01 en el paso "2) Precios de
   proveedores" (#91); seguir viendo Ma01 (#90).
5. Luego Grupos 4 (costeo) y 5 (#66 a #69) del SaaS y el Registro de eventos (lectura): 5 a 7 h.
6. Decisiones del usuario: #83 (menor), #86 (menores) y la primera elección de C8 (Acción Meta Ads) en la copia DEMO.
7. Después: Sesión 3 (Power BI básico y pendientes menores #51, #52, #36, #4, #17, #19, #22, #21) y Sesión 4 (entrega
   final), igual que en la V19. Durante la revisión de los tutores no se hace push (#77).
8. Del usuario: renovar los dos respaldos de RESPALDO_FUENTES antes del **13-oct-2026**.
9. Tiempo programado por el usuario: 5 días para la entrega (3 para lo técnico y 2 para la parte escrita), con chats de
   unas 6 horas.

### Tips aprendidos hoy
- Una migración bien hecha se puede correr dos veces: la segunda responde "Ya estaba aplicada" y no escribe. A eso se le
  llama que sea idempotente.
- "Congelar" con un contador: en vez de guardar la fecha de cada venta, se anota cuántos combos se habían vendido en el
  momento del cambio. Todo lo vendido hasta ese número queda con las unidades viejas; lo que venga después usa las nuevas.
- Un programa puede tener un error que nunca se ve porque todos lo usan de la misma forma (el catálogo por código). Se
  encontró al leer cómo recorre la hoja, no al probarlo.
- El check verde de una prueba dice que el estado final es correcto; para medir tiempos hace falta el aviso o el registro.
- `git status --short` antes de cada commit evita subir por accidente una foto previa o un archivo de datos.

---

# CONTENIDO V21 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V21

*Última actualización: 4-oct-2026 (noche, ~11:00 p. m. CR). Reemplaza a la V20. La V20 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V20", que a su vez trae la V19 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V21".
Subir este archivo y el libro de mapas V23 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.** (Las claves del SaaS están solo en el `.env`, en Render y en
el libro de mapas, Regla #15.)

*Nota de fechas: el chat fue el mismo 4-oct-2026, de las 7:05 p. m. a ~11:00 p. m. (hora CR), a continuación del chat de la
V20. Las horas de las pruebas salen de la Terminal del usuario y de los avisos del SaaS.*

---

## ACTUALIZACIÓN V21 (4-oct-2026, noche, ~4 h) — SaaS Grupo 3a (NODO INTEGRAL y tablas cruce maestras) construido y probado en Render; 'Activación de Promoción' C16 la escribe el robot (#84); decisiones del 3b (combos)

### Estado general
- **Grupo 3a del SaaS CERRADO:** construido, probado en la PC (136 revisiones, TODO OK), en GitHub (**ea7fa98** y
  **6b47b8e**) y probado en Render. Pantallas nuevas (menú "3 Integraciones"): **NODO INTEGRAL** (lectura de las 20 filas
  para los dos usuarios; cuatro celdas manuales solo administración, **#64 resuelto**), **Tabla cruce maestra** y **Tabla
  cruce maestra MP** (solo lectura).
- **#84 nuevo y resuelto — 'NODO INTEGRAL'!C16 "Activación de Promoción":** tenía el texto de relleno "[ON / OFF]" y nadie
  la escribía. Decisión del usuario: **la escribe el robot** (ON = combo activo o descuento general de día de pago / feriado).
  El SaaS solo la muestra ("Oferta hoy: ON / OFF"); **no hay botón** para ella.
- **Regla #40 nueva:** lo ya vendido no se modifica al cambiar un combo o una oferta.
- **Grupo 3b (combos, #65) NO construido:** decisiones tomadas, incluida la composición de los combos (**#85: opción 3**,
  ajustable desde el SaaS aunque el combo ya tenga ventas, con migración m014). Estimado del 3b: 5 a 5.5 h.
- **#57 sigue EN PRUEBA** (la captura de GitHub Actions es del 5-oct) y **#58 en observación**.
- **El Excel maestro NO cambió en este chat** (sin migración). Desde la primera tarea diaria posterior al push, el robot
  escribirá C16 en el maestro. Las escrituras de prueba fueron sobre la copia DEMO.
- Libro de mapas **V23**: Regla #40; #57, #58, #64, #65 y #80 con "|| V23"; nuevos **#84 a #88**.

### Commits de este chat (todos con push)
- **ffc44b7** — bitácora V20 (hecho al cierre del chat anterior).
- **ea7fa98** — "SaaS Grupo 3a: NODO INTEGRAL (celdas manuales #64), lecturas de las tablas cruce maestras y Activacion de
  Promocion C16 escrita por el robot (#84)": 12 files changed, 1299 insertions; 7 archivos nuevos. Push `ffc44b7..ea7fa98`.
- **6b47b8e** — "SaaS NODO INTEGRAL: se quita el boton del indicador de oferta; C16 la escribe solo el robot (#84)":
  5 files changed, 29 insertions, 52 deletions. Push `ea7fa98..6b47b8e`.
- **Último commit de código: 6b47b8e.** Pendiente al cierre: commit de esta bitácora V21 (solo este archivo).

### 1) Revisión de los adjuntos al inicio
- Llegaron los 5 archivos; la captura de GitHub Actions queda para el 5-oct (decisión del usuario: "seguimos hoy mismo").
- `repo_actual.zip`: marca **ffc44b7**, 780,455 bytes, 128 archivos. Bitácora V20 192,697 bytes; mapa V22 286,222.
- Excel maestro: **251,016 bytes**, el mismo de la tarde (39 eventos, último EVT-000039 del 4-oct 12:20 p. m.): la corrida
  de las 12:17 a. m. todavía no había ocurrido. Ningún evento dice "Desde RESPALDO manual" (#58).
- Copia DEMO: **253,288 bytes**, 25 hojas, 47 eventos (último EVT-000047, la anulación de la hornada).

### 2) Grupo 3a — evidencia de celdas (mapa y Excel coinciden)
| Celda | Etiqueta (col. B) | Valor | Quién la lee | En el SaaS |
|---|---|---|---|---|
| C3 | Tarifa Base KM | 500 | 'Interface digital' B4 y el robot (envío) | Editable: colones enteros de 1 a 5,000 |
| C22 | Número SINPE Móvil | 8888-8888 | El bot (catalogo.py) | Editable: ocho dígitos, se guarda 8888-8888 |
| C4 | Meta de Venta Diaria | 150000 | Nadie | Editable de 0 a 99,999,999 (dato de referencia) |
| C8 | Acción Meta Ads | texto "[Subir / Mantener]" | Nadie | Lista Subir / Mantener ("Sin definir" hasta la primera elección) |
| C14, C15 | Ofertas por feriado y día de pago | textos | Nadie | Solo lectura (sin efecto por la Regla #8) |
- Las demás filas las escribe el robot o son fórmulas: lectura. 'TABLA CRUCE MAESTRA' (TablaProductos C6:I15, 8 productos) y
  'TABLA CRUCE MAESTRA MP' (TablaCruceMaestraMP C8:I20, 11 materias primas): el mapa no marca ninguna celda manual.
- **Tres detalles de construcción aceptados por el usuario:** la celda se ubica por su etiqueta de la columna B (como el
  bot); C8 se muestra "Sin definir" y la primera elección reemplaza el relleno (el texto viejo queda en el evento); C3 en
  colones enteros y C4 con tope de seguridad. Un número con separador de miles ("150,000") se rechaza en vez de adivinarlo.

### 3) #84 — 'Activación de Promoción' (C16): la historia completa
- **Hallazgo:** el mapa la describe como "Bloqueado (Lectura) — regla interna" ON / OFF; en el maestro y en la copia tenía
  el relleno "[ON / OFF]". Claude lo reportó como nota para el mapa y el usuario pidió un **ALTO**: "esto debe definirse;
  el SaaS debe indicar si hay oferta o no".
- **Opciones presentadas:** A fórmula del Excel con migración m014 (recomendada por Claude, ~1 h); B solo indicador en el
  SaaS (~20 min); C la escribe el robot.
- **Decisiones del usuario:** (1) **opción C**, el robot; (2) ON = combo activo **o** descuento de día de pago / feriado;
  (3) el indicador solo en la pantalla NODO INTEGRAL; (4) aprobó un botón para la copia DEMO (la tarea diaria no corre sobre
  ella) y, al verlo en Render, lo **quitó**: "la oferta debería ser actualizada por el robot para ser consecuente con el
  mapa de celdas y el proyecto".
- **Cómo quedó:** `activacion_promocion.py` decide con el catálogo compartido (el mismo del bot y del mostrador, que ya
  aplica la Regla #8), escribe solo C16 si el valor cambió y deja el evento **ACTIVACION_PROMOCION** con anterior → nuevo y
  el motivo. `nodo_integral.py` la llama al final (4 renglones): sin migración, porque es un valor en una celda que ya
  existe, como C5 a C7. `tarea_diaria.yml` no cambió.
- **Dos datos que explican el diseño:** (a) en feriado o día de pago la columna A de los combos sigue diciendo ACTIVO (su
  fórmula solo mira la temperatura); quien apaga los combos es el código. Por eso C16 pregunta primero por el descuento
  general y no copia a C20. (b) Con los detonadores actuales (dos combos ≤ 30 °C y dos ≥ 30 °C) siempre hay un combo
  activo: C16 dará ON hasta que cambie la configuración de los combos (3b).
- **Copia DEMO:** C16 ya dice **ON** (lo escribió el robot desde el botón, antes de quitarlo). El robot la revisará al
  guardar un combo (3b). Límite conocido: en la copia el clima y el día de pago no se mueven.
- **Maestro:** C16 pasará de "[ON / OFF]" a "ON" en la primera tarea diaria después del push. **Comprobar el 5-oct.**
- **Idea no aprobada (queda en #86):** que la tarea diaria actualice también NODO INTEGRAL de la copia DEMO (Secret nuevo
  y cambio de `tarea_diaria.yml`, ~45 min).

### 4) Decisiones del usuario (en orden)
1. Seguir el mismo 4-oct; la captura de GitHub (#57) será del 5-oct.
2. "Sí" al 3a con los tres detalles de construcción.
3. ALTO por C16 → opción C (robot), ON = combo o descuento, indicador solo en NODO INTEGRAL.
4. "Sí" al plan de C16 con botón para la copia DEMO.
5. 3b: C y D en solo lectura; detonadores TEMPERATURA y LLUVIA; tope del descuento 30 %.
6. ALTO: **tope del descuento 10 %** ("más realista") y **quitar el botón** (opción A).
7. Composición de los combos: "debe poder hacer el ajuste de sus componentes desde el SaaS, es más práctico"; confirmó que
   lo ya vendido no se modifica (Regla #40); "este cambio sería en un nuevo chat".
8. Ya con los documentos de cierre entregados: **"confirmo la opción: la tercera es la completa"**. Claude preguntó si era
   construirla o solo describirla en el reporte; el usuario eligió **construir la opción 3** (#85) en el chat nuevo.

### 5) Lo construido (tamaños en la PC)
**Nuevos (7):** `microservicios/grupo3_integraciones/nodo_manual.py` (**14048**; `actualizar_dato_manual(campo, valor,
origen)`, `leer_nodo()`), `probar_nodo_manual.py` (**11487**), `activacion_promocion.py` (**7843**; `evaluar()`,
`leer_activacion()`, `actualizar_activacion_promocion(origen)`), `probar_activacion_promocion.py` (**9449**),
`saas_demo/grupo3.py` (**9850**), `plantillas/nodo_integral.html` (**6565**) y `plantillas/tabla_cruce.html` (**2693**).

**Modificados (5):** `nodo_integral.py` (**6333**; 4 renglones al final), `saas_demo/app.py` (**14513**; menú
"3 Integraciones"), `saas_demo/probar_saas_demo.py` (**46221**; secciones L y M), `saas_demo/__init__.py` (906) y
`grupo3_integraciones/__init__.py` (894).

Uso manual: `py -m microservicios.grupo3_integraciones.nodo_manual --local --ver` y
`py -m microservicios.grupo3_integraciones.activacion_promocion --local --ver`.

### 6) Pruebas
- **Entorno de Claude** (copia DEMO de hoy): las tres pruebas TODO OK; el comando de la tarea diaria sobre una copia suelta
  escribió C16 = ON la primera vez y SIN_CAMBIO la segunda (el clima y el tipo de cambio no se consultaron: sin internet).
- **PC del usuario** (9:19 a 9:23 p. m. y 10:23 a 10:27 p. m.): tamaños iguales; `probar_activacion_promocion`,
  `probar_nodo_manual` y `probar_saas_demo` (136 revisiones) con **RESULTADO: TODO OK (0 lineas ERROR)**;
  `datos_prueba\maestro_prueba.xlsx` intacta.
- **Render** (administración, copia DEMO): menú "3 Integraciones"; "Oferta hoy: ON" (combos co01 y co04); tarifa
  500 → 550 → 500, la vuelta con el aviso "Guardado en el Excel en **12.7 s**"; tablas cruce con 8 productos y 11 materias
  primas, sin botones. **Falta la captura de la pantalla ya sin el botón** (despliegue del commit 6b47b8e).
- C8 (Acción Meta Ads) **no se tocó** en la copia DEMO: la primera elección es decisión del usuario.

### 7) Mediciones (dato para el reporte)
- **PC, SaaS Grupo 3a:** tarifa 8.0 y 8.1 s; Acción Meta Ads 8.1 y 8.4 s; activación de promoción 8.0 s; sin cambio 1.0 a
  1.1 s. Prueba del SaaS completa: cerca de 3 min 50 s (17 escrituras).
- **Render Starter, Grupo 3a:** tarifa 12.7 s.
- Tiempo de trabajo (aproximado): lectura y plan ~25 min; 3a ~1 h 15 min (estimado 1 h); C16 ~45 min; quitar el botón ~15 min.

### 8) Grupo 3b (combos) — lo decidido y lo que falta
- **Decidido:** G lista **TEMPERATURA / LLUVIA**; L (descuento) de **0 a 10 %**; C (nombre) en solo lectura, porque es
  llave. D (detalle) primero quedó en solo lectura; con la decisión del #85 ya no se digita: lo escribirá el robot.
- **Propuesto por Claude, sin objeción, POR CONFIRMAR:** H lista <, <=, >, >=; I de 0 a 50 °C con un decimal (con LLUVIA:
  Lluvia / Despejado y el operador no se usa); J de 1 a 9 como dato de referencia; K hasta 150 caracteres; B la reescribe el
  robot; un formulario por combo, solo administración, una escritura y un evento; después el robot revisa C16. Base: 1.5 h.
- **Composición de los combos (#85): DECIDIDO, opción 3, sin construir.** Lo que lleva un combo está en tres lugares que deben coincidir: D5:D8
  (texto), la fórmula de E5:E8 y `TablaComposicionCombos` (D16:D23). El dinero de lo ya vendido está protegido (los precios
  se guardan como número); el inventario no (F = cantidad actual × todos los combos vendidos).

| Opción | Tiempo adicional | Total del 3b |
|---|---|---|
| 1. Composición en solo lectura (trabajo futuro) | 0 | 1.5 h |
| 2. Ajustable con dos componentes; el robot rechaza el cambio si el combo ya tiene ventas (recomendada) | +2 a 2.5 h | ~4 h |
| **3. Ajustable aunque tenga ventas; el robot congela lo ya rebajado (migración m014 en el maestro y en la copia) — ELEGIDA** | +3.5 a 4 h | 5 a 5.5 h |
- Antes de construir, Claude presenta con celdas el diseño de la m014: columnas nuevas en `TablaComposicionCombos` para
  congelar lo ya rebajado; si cambia el producto de un componente, lo ya rebajado sigue cargado al producto anterior; la
  fórmula de E, la tabla y el texto D se escriben en un solo guardado.
- Hallazgos menores (#86): 'DiccionarioDatos' dice que C16 la calcula el Excel; el mapa dice que C19 usa la prioridad (la
  fórmula real junta TODOS los combos activos y nadie lee la columna J); los textos de la columna B están escritos distinto.

### 9) Reglas de trabajo nuevas (van al prompt)
- Los rellenos entre corchetes del Excel ("[ON / OFF]", "[Subir / Mantener]") se presentan al usuario ANTES de construir la
  pantalla que los muestra: pueden ser definiciones pendientes del modelo.
- Una celda que el mapa marca "Bloqueado (Lectura)" no lleva botón en el SaaS: la escribe solo el robot.
- Al leer el mapa 'Tabla combos dinamicos' no se muestra la columna D de las filas 6, 20, 34 y 48 (trae una clave antigua
  de una fuente de clima que ya no se usa); en general se oculta todo texto con "appid=", "key=" o "token=".
- Si en una captura no sale el aviso verde, el estado se comprueba con "Leer de nuevo" y la tabla de la pantalla antes de
  repetir una escritura.

### 10) Libro de mapas V23 (LIBRO_MAPA_DE_CELDAS_PFV23.xlsx, **294425** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- Pestaña nueva **"Cambios V23"** (8 renglones). **Regla #40** nueva; Reglas #37 y #39 con "|| V23".
- Decisiones Pendientes: #57, #58, #64, #65 y #80 con "|| V23"; nuevos **#84** (resuelto: C16), **#85** (decidido, sin
  construir: composición de los combos, opción 3), **#86** (pendientes menores), **#87** (informativo: evidencia y tiempos) y **#88** (lección).
- Bloques "ACTUALIZACIÓN V23" en los mapas 'NODO INTEGRAL' y 'Tabla combos dinamicos'.
- Verificado contra V22 celda por celda: 7,600 iguales, 7 con texto agregado "|| V23", 25 nuevas + pestaña nueva
  (56 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos. 56 pestañas.

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_actual.zip` (780,455; marca ffc44b7), `ANTES_grupo3a_v2.zip` y `ANTES_sin_boton.zip` (carpetas
  saas_demo y microservicios antes de cada paquete), `paquete_grupo3a_v2.zip` (43,726; 12 archivos),
  `paquete_grupo3a_sin_boton.zip` (21,685; 5 archivos). `paquete_grupo3a.zip` (32,874) quedó reemplazado y no se instaló.
- ROBOT PANADERO: sigue `saas_demo.zip` sin seguimiento (se puede mover a download).
- Drive: el maestro no cambió en este chat; la copia DEMO, por las pruebas (C16 = ON; C3 volvió a 500). `.env` y Render: sin cambios.

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V21 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V21.md`).
2. Captura del SaaS en Render ya sin el botón (pantalla NODO INTEGRAL).
3. **#57 y #84 en el maestro:** captura de GitHub Actions del 5-oct (horas de las corridas "Scheduled") y, en el maestro
   recién descargado, el evento ACTIVACION_PROMOCION y C16 = ON. **#58:** seguir observando.
4. **Grupo 3b del SaaS:** confirmar el plan propuesto, presentar el diseño de la migración m014 (#85, opción 3) y construir
   (5 a 5.5 h). Luego Grupos 4 (costeo) y 5 (#66 a #69) y el Registro de eventos. Estimado total: 10 a 12.5 h.
5. Decisiones del usuario: #81 (existencias mínimas), #83 (paso 0 de probar_hornadas, menor), #86 (menores) y la primera
   elección de C8 (Acción Meta Ads) en la copia DEMO.
6. Después: Sesión 3 (Power BI básico y pendientes menores #51, #52, #36, #4, #17, #19, #22, #21) y Sesión 4 (entrega
   final), igual que en la V19. Durante la revisión de los tutores no se hace push (#77).
7. Del usuario: renovar los dos respaldos de RESPALDO_FUENTES antes del **13-oct-2026**.

### Tips aprendidos hoy
- Un mismo dato puede vivir en varios lugares del Excel (texto, fórmula y tabla): cambiarlo en uno solo deja el sistema
  diciendo dos cosas distintas. Por eso esos cambios los hace una sola función del robot, en un solo guardado.
- "Congelar" un valor es guardarlo como número en el momento del hecho (una venta, una compra): así los cambios posteriores
  de precio o de descuento no reescriben la historia.
- El aviso verde del SaaS solo aparece en la página a la que se vuelve después de guardar; al cambiar de página se pierde.
- En `git status --short`, la A es un archivo nuevo y la M uno modificado; los dos signos de pregunta son archivos que Git
  no sigue.

---

# CONTENIDO V20 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V20

*Última actualización: 4-oct-2026 (tarde, ~7:00 p. m. CR). Reemplaza a la V19. La V19 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V19", que a su vez trae la V18 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V20".
Subir este archivo y el libro de mapas V22 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.** (Las claves del SaaS están solo en el `.env`, en Render y en
el libro de mapas, Regla #15.)

*Nota de fechas: el chat fue el 4-oct-2026, de las 3:44 p. m. a ~7:00 p. m. (hora CR). Las horas de las pruebas salen de
los avisos del SaaS, de la Terminal del usuario y de la lista de GitHub Actions.*

---

## ACTUALIZACIÓN V20 (4-oct-2026, tarde, ~3 h 15 min) — SaaS Grupo 2 (Control de inventario y Hornadas) construido y probado en Render; origen SaaS en los eventos; #75 resuelto; prueba de dos horarios de la tarea diaria (#57)

### Estado general
- **Grupo 2 del SaaS CERRADO:** construido, probado en la PC (109 revisiones, TODO OK), en GitHub (**73be99d**) y probado
  en Render. Pantallas nuevas: **Control de inventario** (lectura para los dos usuarios; existencia mínima solo
  administración, **#63 resuelto**) y **Hornadas** (registrar y anular, los dos usuarios).
- **#75 resuelto** (títulos del recibo) y **#77 cerrado por decisión del usuario** (no se hace).
- **Origen de los eventos (#78 nuevo, resuelto):** hornadas, anular venta de despacho e ingreso de inventario ya anotan
  "SaaS (despacho)" o "SaaS (administracion)".
- **#57 comprobado y EN PRUEBA:** la corrida programada volvió a salir a las 12:18 p. m. (6 h 01 min tarde). Decisión del
  usuario: **dos horarios** (12:17 a. m. y 6:17 a. m. CR), commit **fb9d631**. El resultado se lee el 5-oct.
- **#58 sigue en observación** (los eventos del 4-oct leyeron directo).
- **El Excel maestro NO cambió** (sin migración). Todas las escrituras de prueba fueron sobre la copia DEMO.
- **Grupo 3 del SaaS NO empezado:** se leyeron sus mapas y se presentaron hallazgos (#80). **El usuario aprobó el plan
  del 3a (NODO INTEGRAL y lecturas) con sus seis decisiones**; el 3b (combos) sigue sin decidir. Construir el 3a es lo
  primero del próximo chat.
- Libro de mapas **V22**: #57, #58, #63, #75 y #77 con "|| V22"; nuevos **#78 a #83**. No hay Reglas nuevas.

### Commits de este chat (todos con push)
- **81e09ad** — bitácora V19 (hecho al cierre del chat anterior).
- **fb9d631** — "Tarea diaria: dos horarios (00:17 y 6:17 CR) para medir el atraso de GitHub (#57)": 1 file changed,
  14 insertions, 1 deletion. Push `81e09ad..fb9d631`.
- **73be99d** — "SaaS Grupo 2: control de inventario con existencia minima (#63), hornadas, origen SaaS en los eventos y
  titulos del recibo (#75)": 14 files changed, 962 insertions, 14 deletions. Push `fb9d631..73be99d`.
- **Último commit de código: 73be99d.** Pendiente al cierre: commit de esta bitácora V20 (solo este archivo).

### 1) Revisión de los adjuntos al inicio
- Llegaron 3 de 5; después el maestro, el ZIP y la copia DEMO correcta.
- **La primera copia DEMO adjunta era la de ANTES de las pruebas** (249,664 bytes, 143 ventas, último evento
  EVT-000037). La correcta, descargada de Drive (quedó con "(1)" en el nombre): **252,064 bytes**, 25 hojas, 0 errores,
  44 eventos.
- Excel maestro de Drive: **251,016 bytes**; eventos nuevos EVT-000038 (precios OK, 16 de 16) y EVT-000039
  (competencia OK; C21 "Sin promocion vigente (04-10-2026)"), los dos a las 12:20 p. m. del 4-oct.
- `repo_actual.zip`: marca **81e09ad**, 693,094 bytes, 122 archivos, sin `.env` ni `credenciales`.

### 2) Evidencia de las 7 escrituras del Grupo 1 en Render (copia DEMO, 'RegistroEventos')
| Evento | Hora (4-oct) | Tipo | Qué registró | Origen |
|---|---|---|---|---|
| EVT-000038 | 12:38:07 a. m. | VENTA_DESPACHO | Comanda 2 Baguette + 1 CupCake, filas 148 y 149, 2,028.54 | SaaS (despacho) |
| EVT-000039 | 12:51:20 a. m. | ENTREGA_PEDIDO | WA-101, 8,505.55 | SaaS (despacho) |
| EVT-000040 | 1:13:49 a. m. | ANULACION_PEDIDO | WA-118-A, -38,104.04, "prueba en Render" | SaaS (administracion) |
| EVT-000041 | 1:23:42 a. m. | ANULACION | Venta de 1 CupCake, fila 149 | Robot Python |
| EVT-000042 | 1:25:06 a. m. | CONFIG_DESCUENTO_GENERAL | B5 de 5 % a 7 % | SaaS (administracion) |
| EVT-000043 | 1:28:35 a. m. | INGRESO_INVENTARIO | F007, 1 saco de harina, fila 46, costo 17,510.00 | Robot Python |
| EVT-000044 | 1:30:31 a. m. | CONFIG_DESCUENTO_GENERAL | B5 de 7 % a 5 % | SaaS (administracion) |
- Las celdas coinciden: 145 ventas; WA-101 con ENTREGADO = 4-oct 12:51:18 a. m. y ESTADO sin cambio; fila WA-118-A;
  B5 = 5 %; 'Control de Inventario': Baguette 228 → 226, CupCake 123, harina 5.77 → 6.77.
- **Hallazgo:** EVT-000041 y EVT-000043 decían "Robot Python" (funciones viejas sin el dato de quién las usó): de ahí el #78.

### 3) #57 — la corrida programada volvió a salir a las 12:18 p. m.
| Día | Programada | Salió | Atraso |
|---|---|---|---|
| 3-oct (#6) | 7:00 a. m. | 12:18 p. m. | 5 h 18 min |
| 4-oct (#8) | 6:17 a. m. | 12:18 p. m. | 6 h 01 min |
- Corrida #8 "Scheduled", verde, 2 min 27 s. Cambiar el minuto no quitó el atraso; lo que se repite es la HORA de salida.
- `tarea_diaria.yml` está bien (12:17 UTC). En el foro de GitHub hay reportes de 2026 con atrasos de 5 a 7 horas en tareas
  programadas (la corrida manual siempre funciona) y la documentación dice que el horario es de "mejor esfuerzo".
- Dos lecturas posibles: (a) GitHub atrasa ~6 h; (b) GitHub suelta la tarea hacia las 12:18 p. m. pida lo que se pida.
- **Idea del usuario:** mover nuestra hora tomando en cuenta el atraso. **Decisión: opción B = dos horarios**, 12:17 a. m.
  y 6:17 a. m. CR (`17 6 * * *` y `17 12 * * *`). Si (a): salen ~6:17 a. m. y ~12:18 p. m. Si GitHub es puntual:
  12:17 a. m. y 6:17 a. m. (la segunda corrige el clima de madrugada: NODO INTEGRAL usa la temperatura del momento).
  Si las dos salen al mediodía: mover la hora no sirve y queda la opción 1C (disparador externo; clave nueva de GitHub).
- Correr dos veces no duplica datos (mismas celdas); quedan 4 eventos por día y el respaldo se repite.
- Archivo: `tarea_diaria.yml` **5725** bytes (antes 4603). Efecto si sigue saliendo al mediodía: el clima, los combos del
  día y la bandera de día de pago se actualizan a las 12:18 p. m.

### 4) Decisiones del usuario (en orden)
1. #57: dos horarios (opción B), como prueba.
2. Plan del Grupo 2 aprobado: (1) el #63 cubre G5:G22 buscando el producto por nombre; (2) existencia mínima de 0 en
   adelante con hasta 2 decimales; (3) **opción C**: origen "SaaS" también en anular venta e ingreso de inventario;
   (4) #75 **opción C**: cambiar solo los títulos.
3. **#77: no se hace.** "Es agregar más trabajo a algo sencillo; hacer push en la revisión no tiene sentido."
4. La harina queda con mínimo **2** en la copia DEMO.
5. Cierre: eligió "B" (seguir con el Grupo 3) y luego corrigió: **opción A, cerrar el chat con los respaldos**.
6. Ya con los documentos de cierre entregados: **"apruebo el plan del 3a con sus seis decisiones"** (punto 10).

### 5) Lo construido — `paquete_grupo2.zip` (**52605** bytes, 14 archivos; tamaños en la PC)
**Nuevos:**
- `microservicios/grupo2_inventario/existencia_minima.py` (**9214**, #63): `actualizar_existencia_minima(producto,
  minimo, origen)` y `leer_existencias()`. Ubica el producto por su nombre exacto en 'Producto Maestro' y la celda por
  el encabezado 'Stock'; escribe SOLO esa celda; tope de seguridad 99,999; evento CONFIG_EXISTENCIA_MINIMA con valor
  anterior → nuevo; mismo valor = SIN_CAMBIO; nunca escribe sobre una fórmula. Uso manual:
  `py -m microservicios.grupo2_inventario.existencia_minima --local --ver`.
- `probar_existencia_minima.py` (**8852**): sobre copia temporal (nada que restaurar).
- `saas_demo/grupo2.py` (**11637**), `plantillas/control_inventario.html` (**4382**) y `plantillas/hornadas.html` (**5869**).

**Modificados:** `hornadas.py` (**17597**; antes 17010), `ventas_despacho.py` (**7570**; antes 7097),
`ingreso_inventario.py` (**20898**; antes 20588): dato opcional `origen` al final; sin él, "Robot Python" como antes.
`saas_demo/app.py` (**14131**; menú "2 Inventario"), `grupo1.py` (**19480**), `probar_saas_demo.py` (**38049**; secciones
J y K), `plantillas/pedido.html` (**5424**; #75), `saas_demo/__init__.py` (811) y `grupo2_inventario/__init__.py` (517).

**Decisiones de construcción (el usuario las puede cambiar):** tope 99,999 en la existencia mínima; anular una hornada lo
pueden hacer los dos usuarios; solo se anula la última hornada de cada producto (así funciona `anular_hornada`).

### 6) Hallazgos del Grupo 2 (mapa contra Excel)
- El 'Mapa_Control Inventario' describía G5:G11 (7 productos); el Excel tiene `TablaControlInventario` A4:J22, 18
  productos, G = 10 en todos. Con 10 parejo, 15 de 18 salen en REABASTECER. Queda **#81**: decidir los mínimos.
- Ninguna fórmula lee G2 de las hojas de costo: anular una hornada no mueve el modelo.
- Una hornada pequeña encarece el costo unitario (10 Baguette: ₡2,451.03 cada una): en Render se probó con 100.

### 7) Pruebas
- **Entorno de Claude** (copia DEMO del 4-oct): `probar_existencia_minima` TODO OK; `probar_saas_demo` 109 de 109 OK.
  Pruebas viejas sobre copias desechables: ingreso de inventario TODO OK; ventas despacho corrió completa;
  `probar_hornadas` con UNA línea ERROR en su paso 0, que espera una copia sin materia prima (**#83**, no es del cambio).
- **PC del usuario** (5:45 a 5:52 p. m.): los 15 tamaños coinciden; las dos pruebas **RESULTADO: TODO OK (0 lineas ERROR)**;
  `datos_prueba\maestro_prueba.xlsx` intacta.
- **Render** (6:10 a 6:25 p. m., todo como administración):

| Prueba | Resultado | Tiempo |
|---|---|---|
| Hornada de 100 Baguette | Lote H-20261004-1810-P01, costo unitario ₡430.23 | 19.9 s |
| Existencia mínima de la harina 10 → 2 | Celda G13; pasó de REABASTECER a OK (de 15 a 14 productos) | 12.5 s |
| Anular la hornada | Lote ANULA-H-20261004-1810-P01, corrección de -100; harina de vuelta en 6.77 | 12.7 s |
| Recibo de WA-101 (#75) | "Importe con descuento" ₡8,005.55 y "Subtotal sin descuento" ₡8,426.90 | lectura |

### 8) Mediciones (dato para el reporte)
- **PC, SaaS Grupos 1 y 2:** entregar 7.8 s; anular pedido 8.0; comanda 7.7; anular venta 7.8 y 8.3; descuento 8.1 y 7.9;
  factura 8.2 y 8.1; existencia mínima 7.9 y 8.1; **hornada de 10 Baguette 14.3 s**; anular hornada 9.7 s; rechazos 0.9 a 1.1 s.
- **Render Starter, Grupo 2:** hornada 19.9 s; existencia mínima 12.5 s; anular hornada 12.7 s. La hornada tarda más
  porque el robot recalcula dos veces (una para el costo de la hornada y otra al guardar).
- GitHub Actions: atrasos de 5 h 18 min y 6 h 01 min en la corrida programada; corrida #8 de 2 min 27 s.
- Tiempo del Grupo 2: ~1 h 45 min (estimado 1.5 h), incluidos el origen de los eventos y el #75.

### 9) Estado de la copia DEMO al cierre (además de lo anotado en la V19)
- 'Control de Inventario' G13 (harina) = **2**. 'Historico Diario': dos filas nuevas (+100 y -100 Baguette, 4-oct 6:10 p. m.).
- 'costo pan baguette': fecha y hora de esa hornada y G2 = 0 (anulada). Suma neta cero: Baguette 226, harina 6.77.
- 'RegistroEventos': tres eventos nuevos con origen "SaaS (administracion)". Versión 33 de Drive a las 6:16 p. m.

### 10) Hallazgos previos al Grupo 3 (#80) — plan del 3a APROBADO por el usuario; 3b sin decidir
- **NODO INTEGRAL (#64):** de las 6 celdas manuales del mapa, solo C3 (Tarifa Base KM: la leen el robot e 'Interface
  digital' B4) y C22 (SINPE: la lee el bot) tienen efecto. C4 (150000), C8 (texto "[Subir / Mantener]"), C14 (texto
  "[Combo 1") y C15 ("todos los combos") no las lee ninguna fórmula ni el código; C14 y C15 chocan con la Regla #8.
  Propuesta: C3 mayor que 0 con tope 5,000; C22 con formato 8888-8888; C4 editable como dato de referencia; C8 lista
  Subir / Mantener; C14 y C15 solo lectura (opción A) o editables (opción B); solo administración.
- **Tabla Combos Dinamicos (#65):** mapa y Excel coinciden (4 combos, filas 5 a 8). Pero C (nombre) es LLAVE de las
  ventas y del inventario (Regla #16) y D (detalle) es solo texto: las cantidades reales están en la fórmula de E y en
  `TablaComposicionCombos`. Falta decidir detonadores permitidos, tope del descuento y qué hacer con C y D.
- **Orden:** Grupo 3 en dos paquetes: 3a NODO INTEGRAL + lecturas de las tablas cruce maestras (~1 h); 3b combos (~1 h o más).
- **Plan del 3a APROBADO (seis decisiones del usuario):** (1) C3: número mayor que 0, tope 5,000 (el 0 se rechaza);
  (2) C22: ocho dígitos con el formato 8888-8888 (número supuesto, Regla #27); (3) C4: editable, de 0 en adelante, con
  la nota de que hoy es un dato de referencia; (4) C8: lista Subir / Mantener; (5) C14 y C15: SOLO LECTURA, con la nota
  de que la Regla #8 las dejó sin efecto (no se borra nada del Excel, Regla #23); (6) las celdas manuales las cambia
  solo administración y las lecturas las ven los dos usuarios. Incluye la función nueva del #64, la pantalla de NODO
  INTEGRAL y las lecturas de las dos tablas cruce maestras. **No se construyó.** El 3b (#65) sigue sin decidir.

### 11) Reglas de trabajo nuevas (van al prompt)
- Una descarga nueva de Drive queda con "(1)" en el nombre si en download ya hay un archivo igual: esa es la que se adjunta.
- Un archivo suelto se entrega con un nombre único (para que el navegador no le agregue "(1)").
- Antes de cada commit Claude anuncia el resumen esperado ("N files changed, X insertions, Y deletions").
- Si un mensaje llega cortado o hay duda sobre la opción elegida, Claude repite en dos renglones lo que entendió.
- Al leer los mapas, Claude no muestra en pantalla los renglones que traen claves.

### 12) Libro de mapas V22 (LIBRO_MAPA_DE_CELDAS_PFV22.xlsx, **286222** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- Pestaña nueva **"Cambios V22"** (9 renglones). Reglas #37 y #39 con "|| V22".
- Decisiones Pendientes: #57, #58, #63, #75 y #77 con "|| V22"; nuevos **#78** (resuelto: origen del evento), **#79**
  (informativo: evidencia y tiempos), **#80** (hallazgos del Grupo 3: 3a aprobado, 3b pendiente), **#81** (pendiente: existencias
  mínimas), **#82** (lección) y **#83** (pendiente menor: paso 0 de probar_hornadas).
- Bloques "ACTUALIZACIÓN V22" en 'Mapa_Control Inventario' y 'Mapa hoja historico diario'.
- Verificado contra V21 celda por celda: 7,488 iguales, 7 con texto agregado "|| V22", 50 nuevas + pestaña nueva
  (62 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos. 55 pestañas.

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_actual.zip` (693,094; marca 81e09ad), `tarea_diaria_2horas.yml` (5725),
  `tarea_diaria_ANTES_2horas.yml` (4603), `paquete_grupo2.zip` (52,605), `ANTES_grupo2.zip` (carpetas saas_demo y
  microservicios antes del Grupo 2), las dos copias DEMO (la de 249,664 y la nueva de 252,064).
- ROBOT PANADERO: sigue `saas_demo.zip` sin seguimiento (se puede mover a download).
- Drive: el maestro solo cambió por la tarea diaria; la copia DEMO, por las pruebas. `.env` y Render: sin cambios.

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V20 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V20.md`).
2. **#57:** captura de GitHub Actions del 5-oct con las horas de las corridas "Scheduled". **#58:** seguir observando.
3. **Grupo 3 del SaaS:** construir el 3a (plan ya aprobado, punto 10); decidir y construir el 3b (combos). Luego Grupos 4 (costeo) y
   5 (#66 a #69) y el Registro de eventos. Estimado: 6.5 a 8.5 h.
4. Decisiones del usuario: #81 (existencias mínimas) y #83 (paso 0 de probar_hornadas, menor).
5. Después: Sesión 3 (Power BI básico y pendientes menores #51, #52, #36, #4, #17, #19, #22, #21) y Sesión 4 (entrega
   final), igual que en la V19. Durante la revisión de los tutores no se hace push (#77).
6. Del usuario: renovar los dos respaldos de RESPALDO_FUENTES antes del **13-oct-2026**.

### Tips aprendidos hoy
- GitHub entiende las horas en UTC: Costa Rica va 6 horas atrás (12:17 UTC = 6:17 a. m. CR).
- Un mismo archivo de tarea puede tener varios horarios: cada renglón `cron` dispara la tarea completa.
- La llave de sesión es el texto con que el servidor firma el "pase" de quien ingresó; no es la contraseña de nadie.
- `Compress-Archive` guarda carpetas enteras en un ZIP: sirve como respaldo de un solo comando antes de extraer un paquete.
- En Git, la letra A verde es un archivo nuevo ya elegido y la M verde uno modificado ya elegido.

---

# CONTENIDO V19 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V19

*Última actualización: 4-oct-2026 (madrugada, ~1:45 a. m. CR). Reemplaza a la V18. La V18 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V18", que a su vez trae la V17 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V19".
Subir este archivo y el libro de mapas V21 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.** (Las dos claves del SaaS están solo en el `.env`, en Render y en
el libro de mapas V21, Regla #15.)

*Nota de fechas: el chat fue del 3-oct-2026 ~7:40 p. m. al 4-oct-2026 ~1:45 a. m. (hora CR), unos minutos después de
cerrar la V18. Las horas son de CR; las de Render salen de sus registros (que muestran hora CR) y las del SaaS, del aviso
de cada pantalla. **No hubo cambios de código en este chat:** todo fue configuración en Drive, en el `.env` y en Render.*

---

## ACTUALIZACIÓN V19 (3-oct noche a 4-oct-2026 madrugada, ~6 h) — El SaaS de demostración YA ESTÁ EN RENDER: copia DEMO en Drive, servicio `saas-panaderia` en Starter y Grupo 1 probado en la nube con los dos usuarios

### Estado general
- **El SaaS de demostración funciona en internet:** `https://saas-panaderia-um7r.onrender.com` (servicio nuevo de Render
  `saas-panaderia`, plan Starter, aparte del bot). Trabaja SOLO sobre la copia `Estudio_Financiero_Panadero_DEMO.xlsx`
  de Drive. **Regla #39 cumplida en la nube** (los tres candados pasan).
- **Grupo 1 probado completo en Render** con los dos usuarios: 7 escrituras, todas correctas, entre **11.7 y 12.3 s**
  (en la PC del usuario: 7.8 a 8.5 s). Los permisos por usuario funcionan.
- **El Excel maestro NO cambió** (sin migración; el SaaS no lo conoce). El bot y la tarea diaria no se tocaron.
- **Sin commits de código.** Último commit de código: **0de7393**. `b7027c1` = bitácora V18. Pendiente al cierre: commit
  de esta bitácora V19 (solo este archivo).
- **#57 sigue pendiente** (al cierre, 1:45 a. m. del 4-oct, la corrida de las 6:17 a. m. aún no había ocurrido) y **#58
  sigue en observación** (0 de 37 eventos dicen "Desde RESPALDO manual").
- Decisiones Pendientes nuevas en el libro de mapas V21: **#73** (resuelto: SaaS en Render), **#74** (lección), **#75**
  (pendiente menor: presentación del recibo), **#76** (informativo: tiempos y estado de la copia DEMO) y **#77**
  (pendiente opcional: llave de sesión fija). No hay Reglas nuevas.

### 1) Revisión de los adjuntos al inicio
- Bitácora V18 (159,389 bytes) y libro de mapas V20 (272,583 bytes, 53 pestañas): los vigentes.
- Excel maestro de Drive: 250,884 bytes, 25 hojas, 0 errores; último evento EVT-000037 (3-oct, 3:00:05 p. m.). Es el
  mismo estado que revisó la V18 (nada escribió en Drive entre los dos chats).
- **El primer `repo_actual.zip` era el VIEJO** (marca 7054d12, 522,671 bytes, sin `saas_demo`): se repitió la lección
  #72. El segundo trajo la marca **b7027c1** con los 18 archivos de `saas_demo`, idénticos a los de la PC (el usuario
  adjuntó además `saas_demo.zip` con su carpeta: 18 archivos, 0 diferencias). Dentro del ZIP los tamaños son mayores
  que en la bitácora (p. ej. `config.py` 6,871 contra 6,727) porque Git en Windows guarda cada fin de renglón con un
  carácter más; el contenido es el mismo (159,389 + 1,966 renglones = 161,355).
- **#57:** no comprobable (el chat empezó y terminó antes de las 6:17 a. m. del 4-oct). **#58:** sin eventos de respaldo.
- 'Interface digital': 22 pedidos; solo 4 con TELEFONO (WA-119 a WA-122, filas 26 a 29), todos del celular del usuario.

### 2) Decisiones del usuario (en orden)
1. **Teléfonos de la copia DEMO: opción C** = se cambian por un número supuesto (Regla #27). El nombre de perfil de
   WhatsApp queda en la columna CLIENTE.
2. Orden confirmado: copia DEMO en Drive → servicio en Render (Starter) → prueba con los dos usuarios → Grupo 2.
3. La copia se sube **sola** a la carpeta del maestro (sin carpeta propia), igual que el respaldo y el cache.
4. **Las dos claves del SaaS las eligió el usuario** (las cambió una vez por unas "más amigables"). Pidió documentarlas
   en el libro de mapas que no va a GitHub: están en el **libro V21, Regla #39, bloque "ACTUALIZACIÓN V21"**.
5. Hacer esa misma madrugada las pruebas de escritura en Render: primero tres y después las tres operaciones que
   faltaban del Grupo 1. Luego, el cierre (esta bitácora, el mapa V21 y el prompt). El Grupo 2 pasa al chat siguiente.

### 3) La copia de demostración (Regla #39, candado 3)
- Archivo: **`Estudio_Financiero_Panadero_DEMO.xlsx`**, 249,664 bytes. Es el maestro descargado de Drive el 3-oct (último
  evento EVT-000037) con UN solo cambio: 'Interface digital' **P26:P29** (TELEFONO de WA-119 a WA-122) =
  **+50688887777**, el número supuesto de las pruebas del robot.
- **Cómo se hizo:** sin abrir ni volver a guardar el libro con ningún programa. Dentro del .xlsx el teléfono era UN
  texto compartido usado por 4 celdas; se cambió solo ese texto. De las 76 partes internas del archivo cambió una (la
  de los textos). Pesa menos que el maestro (250,884) por el empaque, no por el contenido.
- **Doble revisión:** (a) `py -m migraciones.comparar_libros`: 25 hojas, 5,059 fórmulas iguales, 4 celdas distintas
  (P26 a P29), 0 valores calculados cambiados, 0 errores; (b) ensayo del recálculo del robot con LibreOffice: 2.5 s,
  0 errores, saldo acumulado 825,048, Baguette 228 y CupCake 123 sin cambio.
- **Por qué no afecta las pruebas:** el bot trabaja sobre el maestro y contesta al número que Twilio le entrega en cada
  mensaje; el SaaS no envía WhatsApp (no tiene las claves de Twilio): solo muestra "termina en 7777" y el recibo en
  pantalla. Los pedidos nuevos del bot entran al maestro, no a la copia.
- **En Drive (3-oct, 9:52 p. m.):** subida a "Mi unidad" junto al maestro, 244 KB, icono de Excel, compartida con la
  cuenta de servicio como Editor. **ID de la copia: `1L62bTSNFMcRG4n5IRxP0s59Xdg_m1VKx`** (distinto del maestro:
  candado 2). NUNCA abrirla con Google Sheets (Regla #12).

### 4) El `.env` de la PC
- Tres renglones nuevos al final (39 a 41): `SAAS_CLAVE_DESPACHO`, `SAAS_CLAVE_ADMIN` y `SAAS_EXCEL_FILE_ID` (el ID de
  la copia). Las dos claves deben ser distintas, de 8 caracteres o más, o el SaaS no arranca (candado 1).
- **Comprobación desde la PC** (carpeta ROBOT PANADERO):
  `py -c "from saas_demo import datos; print(datos.verificar_copia(reciente=True))"` → respondió
  `Estudio_Financiero_Panadero_DEMO.xlsx`. Prueba los tres candados sin escribir nada y sin mostrar claves.
- Un `#` al final de una clave se lee bien desde el `.env` (se comprobó: forma parte del valor si no lleva espacio antes).

### 5) El servicio del SaaS en Render
- **+ New → Web Service → repositorio Gigo1965/robot-panaderia.** Nombre **`saas-panaderia`**, Docker, rama `main`,
  región Virginia (US East), plan **Starter** (US$7/mes; con el bot son US$14/mes mientras existan los dos).
  ID del servicio `srv-db0ub1ou01pc73bvv0t0`. URL **https://saas-panaderia-um7r.onrender.com** (Render agregó `-um7r`).
- **Health Check Path:** `/salud`. **Docker Command** (reemplaza el CMD del Dockerfile, que arranca el bot):
  `gunicorn --workers 1 --threads 4 --timeout 120 --bind 0.0.0.0:10000 saas_demo.app:app`
  (puerto fijo 10000 para no depender de que Render expanda `$PORT` en ese campo).
- **Variables (solo 4):** `SAAS_EXCEL_FILE_ID`, `GOOGLE_SERVICE_ACCOUNT_FILE` = `/etc/secrets/clave_google.json`,
  `SAAS_CLAVE_DESPACHO` y `SAAS_CLAVE_ADMIN`. **Secret File** `clave_google.json` = contenido de la llave JSON de la
  cuenta de servicio. **No lleva** `EXCEL_MAESTRO_FILE_ID` ni variables de Twilio. No se usaron "Add from .env" (habría
  metido todo el `.env`) ni "Generate".
- **Antes de crear el servicio** se simuló el arranque solo con esas variables: `/salud` respondió 200 y el robot quedó
  apuntando al ID de la copia.
- **Primer despliegue:** 3-oct, 11:35 p. m., **1 min 12 s**, commit b7027c1, gunicorn 26.2.0 escuchando en 0.0.0.0:10000.
- Redespliegues por cambios de configuración: 11:49 p. m. (Secret File), uno más por las claves nuevas (hora no
  anotada), 12:17 a. m. (nombre del Secret File) y 12:29 a. m. (contenido del Secret File).

### 6) Hallazgo — tres tropiezos de digitación en Render (ninguno del código; ~57 min del servicio creado a la primera lectura)
1. **"El usuario o la contraseña no coinciden."** La clave guardada en Render no era igual a la que se tecleaba. El
   usuario eligió claves nuevas y las cambió en el `.env` y en Render.
2. **"No encontré la clave de la Cuenta de Servicio … /etc/secrets/clave_google.json".** El Secret File se había
   guardado como `clave_google.jsom` (una letra). Se corrigió el nombre.
3. **"Error de Google Drive: Expecting value: line 1 column 1 (char 0)".** En el contenido del Secret File estaba pegado
   el NOMBRE del archivo de la llave y no su contenido (ese mensaje exacto se reprodujo con un texto que no empieza con
   `{`). Se volvió a copiar con la Terminal comprobando el largo del portapapeles (**2387**) y se pegó con Ctrl+A, Ctrl+V.
- **Dato para el reporte:** en los tres casos el SaaS dijo con claridad qué faltaba; los candados y los mensajes
  hicieron el diagnóstico rápido. Primera lectura correcta: **4-oct, 12:32 a. m.** (22 pedidos: 20 por entregar,
  2 entregados, 0 anulados).

### 7) Pruebas en Render (4-oct-2026, madrugada; todas sobre la copia DEMO)
| # | Hora | Usuario | Operación | Resultado | Tiempo |
|---|---|---|---|---|---|
| 1 | 12:38 a. m. | despacho | Venta de despacho: comanda de 2 Baguette + 1 CupCake | Total ₡2,028.54; 'Últimas ventas' 143 → 145 | 11.9 s |
| 2 | 12:51 a. m. | despacho | Marcar entregado WA-101 | Recibo en pantalla, total ₡8,505.55 | 11.7 s |
| 3 | 1:13 a. m. | administración | Anular WA-118 (motivo "prueba en Render") | Fila WA-118-A por ₡-38,104.04 | 11.9 s |
| 4 | entre 1:13 y 1:25 a. m. | administración | Anular la venta de 1 CupCake | Estado Anulada; siguen 145 registros | 11.9 s |
| 5 | 1:25 a. m. | administración | Descuento de día de pago 5 % → 7 % | "Hoy está en 7 %" | 11.7 s |
| 6 | entre 1:25 y 1:30 a. m. | administración | Factura F007, Solera, 1 x HARINA FLORES 25 KG FUERTE | ₡17,510.00, costo de la TCM congelado | 12.3 s |
| 7 | 1:30 a. m. | administración | Descuento 7 % → 5 % (se dejó como el maestro) | "Hoy está en 5 %" | 12.1 s |
- **Permisos comprobados:** `despacho` ve solo Ventas despacho y Pedidos digitales, y en la ficha de un pedido lee "Solo
  administración puede anular un pedido"; `administracion` ve además Descuento de día de pago e Ingreso de inventario.
- Después de la entrega: Por entregar 19, Entregados 3, Anulados 0, Todos 22. Después de la anulación: Anulados 1.
- La copia iba en la **versión 18 de Drive** a la 1:25 a. m. y en la **24** a la 1:30 a. m.
- Cada escritura deja su evento en 'RegistroEventos' de la copia (origen SaaS). **No se descargó la copia para leerlos:**
  queda como comprobación del próximo chat.

### 8) Mediciones (dato para el reporte)
- **SaaS en Render Starter (0.5 CPU, 512 MB): 11.7 a 12.3 s por escritura** (7 mediciones; promedio 11.9 s). En la PC:
  7.8 a 8.5 s. La diferencia (unos 4 s) es coherente con el medio procesador; el tiempo es parejo.
- Despliegue del servicio nuevo: 1 min 12 s. Largo de la llave de Google: 2387 caracteres.
- Recálculo de la copia con LibreOffice en el entorno de Claude: 2.5 s.

### 9) Estado de la copia DEMO al cierre (difiere del maestro en esto)
- 'Interface digital': P26:P29 con el número supuesto; WA-101 con hora en ENTREGADO; fila de anulación WA-118-A; B5 = 5 %.
- 'Ventas despacho': 2 filas nuevas del 4-oct, 12:38 a. m. (Baguette 2, correcta; CupCake 1, anulada).
- 'Ingreso de Inventario': factura F007 (harina, 1 saco). 'RegistroEventos': los eventos de las 7 escrituras.
- El maestro sigue con último evento EVT-000037 (más lo que escriba la tarea diaria del 4-oct).

### 10) Observaciones que quedan anotadas
- **#75 (menor, presentación):** en el recibo de un pedido con descuento, el renglón del producto sale ya con el
  descuento (WA-101: ₡8,005.55) y debajo el "Subtotal" sale sin él (₡8,426.90), luego Descuento -₡421.35 y Envío ₡500.
  Las cuentas son correctas (total ₡8,505.55); la presentación puede confundir a un tutor. Decide el usuario.
- **#77 (opcional):** cada reinicio o despliegue del servicio cierra las sesiones, porque la llave de sesión se crea al
  arrancar. Con la variable `SAAS_SECRET_KEY` fija en Render eso no pasaría. No hay otro efecto. Decide el usuario.
- **Para el Grupo 2 (#63), evidencia ya vista:** el 'Mapa_Control Inventario' describe la existencia mínima en G5:G11
  (extracto de 7 productos) y el Excel tiene `TablaControlInventario` A4:J22 con 18 productos y G5:G22 = 10 en todos.
  Se presenta al usuario con el plan del Grupo 2, antes de construir.

### 11) Reglas de trabajo nuevas (van al prompt)
- Para pegar una llave en una página: copiarla con la Terminal comprobando el **largo** del portapapeles, y el mensaje
  con los pasos de la página **no lleva recuadros** (para no copiar otra cosa por error). Pegar con Ctrl+A, Ctrl+V.
- Los nombres que se escriben en Render (variables, Secret File) se copian con el botón del recuadro, no a mano.
- En las pruebas del SaaS la captura debe incluir el aviso verde ("Guardado en el Excel en N s").
- Después de cada despliegue del SaaS hay que ingresar de nuevo. Cinco intentos fallidos bloquean el ingreso 5 minutos.
- Las claves del SaaS no se guardan en la memoria de Claude ni en la bitácora: `.env`, Render y libro de mapas.

### 12) Libro de mapas V21 (LIBRO_MAPA_DE_CELDAS_PFV21.xlsx, **278434** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- Pestaña nueva **"Cambios V21"** (9 renglones).
- **Reglas Globales:** Regla #39 con "|| V21" y bloque nuevo **"ACTUALIZACIÓN V21"** con los datos del servicio de
  Render, el ID de la copia DEMO, los dos usuarios **y sus claves**.
- **Decisiones Pendientes:** #57, #58 y #63 con "|| V21"; nuevos #73 (resuelto), #74 (lección), #75 (pendiente menor),
  #76 (informativo) y #77 (pendiente opcional).
- Nada del V20 se borró.

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_actual.zip` nuevo (marca b7027c1; reemplazó al de 522,671 bytes),
  `Estudio_Financiero_Panadero_DEMO.xlsx` (249,664 bytes; la misma que está en Drive, antes de las pruebas).
- Drive: archivo nuevo `Estudio_Financiero_Panadero_DEMO.xlsx` junto al maestro. El maestro, el cache, el respaldo
  diario y RESPALDO_FUENTES no se tocaron.
- `.env` de la PC: tres renglones `SAAS_` nuevos. Render: servicio nuevo `saas-panaderia`; el del bot, sin cambios.

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V19 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V19.md`).
2. **#57** (hora real de la corrida "Scheduled" del 4-oct, cerca de las 6:17 a. m.) y **#58** (evento "Desde RESPALDO
   manual"): se revisan con el Excel maestro recién descargado y la lista de corridas de GitHub Actions.
3. **Grupo 2 del SaaS:** Control de inventario (lectura y #63) y Hornadas. Primero el plan con las celdas como evidencia
   (punto 10). Luego Grupos 3 (#64, #65), 4 (costeo) y 5 (#66 a #69) y el Registro de eventos. Estimado: 8 a 10 h.
4. Decisiones del usuario: #75 (presentación del recibo) y #77 (llave de sesión fija).
5. Leer los eventos de la copia DEMO (descargarla de Drive) para dejar la evidencia de las 7 escrituras.
6. Después: Sesión 3 (Power BI básico y pendientes menores #51, #52, #36, #4, #17, #19, #22, #21) y Sesión 4 (entrega
   final: a los tutores se les da la URL del SaaS, los dos usuarios y sus claves), igual que en la V18.
7. Del usuario: renovar los dos respaldos de RESPALDO_FUENTES antes del **13-oct-2026**.

### Tips aprendidos hoy
- Un "Secret File" de Render es un archivo que el servicio encuentra en `/etc/secrets/<nombre>`: el nombre debe ser
  exacto, y el contenido es lo que el archivo lleva adentro, no su nombre.
- El campo "Docker Command" de Render cambia con qué programa arranca la misma caja de Docker: por eso un solo
  repositorio da dos servicios (el bot y el SaaS).
- El ID de un archivo de Drive no es una clave: solo lo abren las cuentas con las que se compartió.
- Un archivo .xlsx es un paquete de partes: se puede cambiar un texto sin recalcular ni tocar lo demás.
- "Live" en Render solo dice que el programa arrancó; que lea el Excel se comprueba entrando al SaaS.

---

# CONTENIDO V18 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V18

*Última actualización: 3-oct-2026 (noche, ~7:30 p. m. CR). Reemplaza a la V17. La V17 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V17", que a su vez trae la V16 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V18".
Subir este archivo y el libro de mapas V20 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: el chat fue el 3-oct-2026, de ~4:05 p. m. a ~7:30 p. m. (hora CR), minutos después de cerrar la V17.
Cierre: el usuario eligió primero respaldar (bitácora y prompt) y después decidió construir también el libro de mapas V20
y pasar a un chat nuevo. **El SaaS en Render NO se empezó: es lo primero del próximo chat.***

---

## ACTUALIZACIÓN V18 (3-oct-2026, tarde y noche) — Sesión 2: SaaS de demostración por módulos. Base + Grupo 1 completo, probado en la PC y en GitHub

### Estado general
- **Existe el SaaS de demostración** (carpeta nueva `saas_demo\`): ingreso con dos usuarios, menú por grupos en el
  orden del Excel maestro y **Grupo 1 (Captura operativa) completo**: Ventas despacho (comanda con total a cobrar),
  Pedidos digitales (entregar y anular), descuento de día de pago y feriado (B5) e Ingreso de inventario.
- **Probado en la PC del usuario: RESULTADO TODO OK** (79 revisiones) sobre una copia temporal. En GitHub: commit
  **0de7393**. **Todavía NO está en la nube**: falta la copia de demostración en Drive y el servicio nuevo en Render.
- **Cambio de alcance (decisión del usuario):** el SaaS no es solo "entregar y anular": debe tener **un módulo por cada
  microservicio**, con todos los valores que los mapas de celdas marcan como digitados a mano. Los mapas son la guía;
  no se improvisa. Estimación del SaaS completo: 11 a 13 h en total (faltan 8 a 10 h, en dos o tres chats).
- **Regla #39 nueva** y **Decisiones Pendientes #62 a #72** (libro de mapas V20; punto 5). Resueltos hoy: **#54, #62 y
  #70**. Pendientes: #63 a #69. #71 informativo y #72 lección.
- El Excel maestro **no cambió** (sin migración, sin escrituras en Drive en este chat). El bot y la tarea diaria no se
  tocaron. Ningún archivo del bot ni de los microservicios existentes fue modificado: solo se AGREGARON archivos.
- #57 y #58 de la V17 siguen pendientes (el chat fue el mismo 3-oct; la corrida de las 6:17 a. m. se estrena el 4-oct).

### Commits de este chat (todos con push)
- **7054d12** — bitácora V17 (hecho al cierre del chat anterior; push `005bf30..7054d12`).
- **cbceaca** — "SaaS de demostracion: base con ingreso, pedidos digitales, entregar y anular (#54)": 14 files changed,
  1562 insertions (13 archivos nuevos de `saas_demo` + `.env.example`).
- **0de7393** — "SaaS Grupo 1: comanda de despacho (#70), descuento de dia de pago (#62), ingreso de inventario y menu
  por grupos": 12 files changed, 1200 insertions, 57 deletions. Push `cbceaca..0de7393  main -> main`.
- **Último commit de código: 0de7393.** Pendiente al cierre: commit de esta bitácora V18 (solo este archivo).

### 1) Revisión del Excel de Drive adjuntado al inicio (250,884 bytes, 25 hojas, 0 errores)
- Último evento: EVT-000037, 3-oct 3:00:05 p. m. (corrida #7). Eventos del 3-oct: EVT-032/033 (10:44 a. m., #5, los dos
  PARCIAL), EVT-034 PARCIAL (Solera) y EVT-035 OK (Musi) (12:23 p. m., #6), EVT-036/037 OK (3:00 p. m., #7, directo).
- **#58:** ningún evento dice "Desde RESPALDO manual". Sigue en observación.
- **#57:** no se pudo comprobar (el horario de las 6:17 a. m. se estrena el 4-oct).
- WA-121 y WA-122 siguen CONFIRMADO sin entregar; NODO INTEGRAL C21 "Sin promocion vigente (03-10-2026)"; 'Flujo caja
  diario' fila 220 (3-oct) sin ventas, saldo acumulado 825,048.00.

### 2) Hallazgo de arquitectura que decidió dónde vive el SaaS
- **El robot conoce UN solo Excel:** `conexion_drive/config.py` lee `EXCEL_MAESTRO_FILE_ID` una vez, al arrancar, y
  todos los microservicios escriben ahí. Las funciones `marcar_entregado` y `anular_pedido_digital` ya existían.
- Opciones presentadas: **A** segundo servicio en Render con el ID de la copia; **B** páginas dentro del servicio del
  bot (habría que cambiar `drive_excel.py` para manejar dos Excel en un mismo programa; el bot y la demostración
  competirían por la misma máquina y por LibreOffice; cada despliegue reinicia el bot); **C** solo en la PC.
- **Decisión del usuario: opción A y pasar ese servicio a Starter** (otros US$7 al mes mientras esté activo).

### 3) Decisiones del usuario (en orden)
1. Opción A (segundo servicio en Render) en plan Starter.
2. **Dos usuarios:** `despacho` y `administracion`.
3. "Si se ha hecho todo este proyecto tan completo no tiene sentido que el SaaS no incluya como módulo cada uno de los
   microservicios." El menú sigue el orden del Excel maestro y del libro de mapas (los 5 grupos).
4. "Hay que seguir las indicaciones de los mapas de celdas, son la guía, sin improvisar." Las funciones que falten se
   anotan como pendientes. "El tiempo no veo problema si queda como debe de ser."
5. Permisos: **despacho** hace ventas de despacho, entregas y hornadas, y ve las lecturas; **administración** hace
   todo, incluidas anulaciones de pedidos digitales, compras y costeo.
6. Subir a Render al cerrar el Grupo 1.
7. Anulación de una **venta de despacho**: la hace el despachador (es el empleado de la tienda; consulta al gerente
   antes). Anulación de un **pedido digital**: solo el gerente (administración).
8. Descuento de día de pago y feriado (B5): **máximo 10 %**; es sobre el precio de venta, los días que hay descuento,
   y no aplica a los productos dentro de un combo (el combo tiene precio especial).
9. Ventas despacho: **opción B = comanda en pantalla con total a cobrar** (#70), sin cambiar la estructura del Excel.
10. Cierre: primero bitácora V18 y prompt como respaldo (opción C); después: "mejor crear ya el libro y migro a otro chat
    nuevo". Render pasa al próximo chat.

### 4) Lo construido (tamaños en la PC del usuario)
**Carpeta `saas_demo\` (18 archivos):**
- `config.py` (**6727**): se importa ANTES que el robot. Lee `SAAS_EXCEL_FILE_ID` (ID o enlace de la copia) y se lo
  entrega al robot como "su" Excel. **Tres candados:** (1) sin esa variable el SaaS no arranca; (2) si es el mismo ID
  que `EXCEL_MAESTRO_FILE_ID`, no arranca; (3) antes de leer o escribir, el nombre del archivo de Drive debe contener
  **DEMO**. Claves de los dos usuarios desde variables (mínimo 8 caracteres, distintas).
- `datos.py` (**11080**): candado 3, lectura del Excel con cache de 30 s, lista de pedidos, y `escribir()`: una
  escritura a la vez, con el tiempo medido. `seguridad.py` (**2265**): sesión, permisos y envío único de formularios.
- `app.py` (**13842**): ingreso, bloqueo tras 5 intentos fallidos, ficha contra formularios falsos, menú, Pedidos
  digitales (`marcar_entregado`, `anular_pedido_digital` solo administración = **#54 resuelto**), `/salud`.
- `grupo1.py` (**19323**): Ventas despacho, descuento general e Ingreso de inventario.
- `ensayo.py` (**2504**): `py -m saas_demo.ensayo` arranca el SaaS en la PC sobre una **copia temporal** de
  `datos_prueba\maestro_prueba.xlsx` (http://127.0.0.1:5001); no toca Drive ni la copia local; no hay que restaurar.
- `probar_saas_demo.py` (**27413**): prueba automática sobre copia temporal (79 revisiones, unos 2 minutos).
- `__init__.py` (579); plantillas: `base.html` (3194), `entrar.html` (1053), `pedidos.html` (3196), `pedido.html`
  (5300), `problema.html` (311), `ventas_despacho.html` (6000), `descuento_general.html` (1275),
  `ingreso_inventario.html` (4176); `estaticos\estilo.css` (13980) y `pantalla.js` (947).
- `.env.example` (**2346**; antes 1748): documenta `SAAS_EXCEL_FILE_ID`, `SAAS_CLAVE_DESPACHO`, `SAAS_CLAVE_ADMIN`.

**Microservicios nuevos (carpeta `microservicios\grupo1_captura_operativa\`):**
- `descuento_general.py` (**5584**, #62): `actualizar_descuento_general(porcentaje)` escribe 'Interface digital' B5
  (ubicada por su etiqueta), acepta de 0 a 10 %, deja evento CONFIG_DESCUENTO_GENERAL con valor anterior -> nuevo.
- `comanda_despacho.py` (**6717**, #70): `cotizar_comanda_despacho` y `registrar_comanda_despacho`: varias líneas en
  un solo guardado, una fila por producto en Tabla6 con la misma fecha y hora, un evento VENTA_DESPACHO por comanda con
  el total. Usa las mismas reglas de precio de `ventas_despacho.py` (que NO se modificó). Si el total cambió entre lo
  que vio el despachador y el registro, rechaza sin escribir. La anulación sigue siendo producto por producto.

**Decisiones de construcción (el usuario las puede cambiar):** la pantalla de compras rechaza la producción propia
(entra por la hornada, Regla #34); el proveedor es obligatorio; el recibo de una entrega se muestra en pantalla (el
SaaS no envía WhatsApp); el teléfono del cliente se muestra oculto ("termina en 1234").

### 5) Valores digitados a mano según los mapas y pendientes #62 a #72 (ya anotados en el libro V20)

| Grupo | Hoja | Lo que digita una persona | Estado |
|---|---|---|---|
| 1 | Ventas despacho | Producto (lista) y cantidad; anulación | **Hecho** (#70 comanda) |
| 1 | Interface digital | B5 % de descuento día de pago / feriado; entregar y anular | **Hecho** (#62, #54) |
| 1 | Detalle comanda digital | Nada (lo escribe el robot) | — |
| 1 | Ingreso de Inventario | Fecha, factura, proveedor, producto y cantidad | **Hecho** |
| 2 | Control de Inventario | G: existencia mínima por producto | **#63** pendiente |
| 2 | Historico Diario | Hornada: producto y cantidad; anulación | Función existe; falta pantalla |
| 3 | NODO INTEGRAL | C3 tarifa km, C4 meta de venta, C8 acción Meta Ads, C14, C15, C22 SINPE | **#64** pendiente |
| 3 | Tablas cruce maestras | Nada (precios de las API de catálogo) | Lectura |
| 3 | Tabla Combos Dinamicos | Por combo: nombre, detalle, detonador, operador, límite, prioridad, mensaje, descuento | **#65** pendiente |
| 4 | Costo baguette / cupcakes | Precio de referencia | Función existe; falta pantalla |
| 4 | Planilla / Gasto operacional / Valor activos | Personas y sueldo; 9 rubros por mes; activos | Funciones existen; faltan pantallas |
| 5 | Flujo caja diario | G: gastos, vales y caja chica del día | **#66** pendiente |
| 5 | Flujo caja anual | B19:M19 salidas de efectivo por mes | **#67** pendiente |
| 5 | Flujo caja proyecto | C28 tasa, B31 plazo, parte de B37:B42 | **#68** pendiente |
| 5 | Punto de Equilibrio | B11 mezcla, B22 % utilidad (verificar B22/B23 contra el Excel) | **#69** pendiente |
| 5 | TIR VAR y proyecciones | Nada (100 % automáticas) | Lectura |

- **#71** (informativo): hallazgo de los productos externos con el 5 %, tiempos del SaaS en la PC y pedidos SINPE.
  **#72** (lección): decimales en las pruebas, marca del commit en el ZIP, avisar cuando un paquete reemplaza a otro.
- Cada función nueva va en la carpeta de su microservicio (no en el SaaS), con su prueba.
- Orden de construcción aprobado: Grupo 1 (hecho) -> Render -> Grupos 2, 3, 4, 5 y Registro de eventos (lectura).

### 6) Pruebas
- **Entorno de Claude** (copia del Excel de Drive del 3-oct, LibreOffice): 79 de 79 OK.
- **PC del usuario:** primera versión (base) TODO OK; Grupo 1: una línea ERROR que era de la PRUEBA, no del sistema
  (para la computadora 8.55 - 1 = 7.550000000000001; se comparaba sin redondear). Corregida la prueba (27413 bytes):
  **RESULTADO: TODO OK (0 lineas ERROR)**, y `datos_prueba\maestro_prueba.xlsx` quedó igual (misma huella).
- Faltaba la librería Flask en la PC: `py -m pip install -r whatsapp_bot/requirements.txt` (instaló flask 3.1.3 y twilio).
- `py -m saas_demo.ensayo` arrancó bien en la PC (http://127.0.0.1:5001).

### 7) Mediciones en la PC (dato para el reporte; copia local, LibreOffice de la PC)
- Cada escritura: **7.8 a 8.5 s** (entregar 8.1; anular pedido 7.9 y 8.1; comanda de 3 líneas 7.8 y 8.5; anular venta
  7.8 a 8.2; descuento 7.8 y 8.1; factura de compra 7.8 y 8.5). La primera escritura tras encender: 12.5 s.
- Un rechazo (no escribe): **0.8 a 1.0 s**.
- Duración de este tramo: de ~4:05 p. m. a ~7:00 p. m. (unas 3 h) para la base y el Grupo 1 completo.

### 8) Hallazgos
- **Pedidos SINPE no se anulan (Regla #30):** WA-121 y WA-122 son SINPE; la anulación se prueba con un pedido en
  efectivo (p. ej. WA-118) y los SINPE sirven para mostrar el rechazo.
- **Productos de compra externa con el descuento del 5 %** (dato de la TCM: precio = costo x 1.05): los días de
  descuento quedan apenas bajo el costo (leche 1,047.38 contra 1,050). **Decisión del usuario: queda como está**
  (el descuento es sobre el precio de venta). Límites: Baguette queda bajo costo desde 48.9 %, CupCake desde 27.7 %.
- La hoja 'Ventas despacho' no tiene número de comanda ni total por cliente (eso solo existe en 'Interface digital'):
  de ahí el #70.
- Las ventas de despacho de la función vieja quedan con origen "Robot Python"; la comanda nueva escribe
  "SaaS (despacho)" o "SaaS (administracion)".
- El primer `repo_actual.zip` adjuntado era el viejo (commit f9b7535): se detectó por la marca del commit dentro del
  ZIP. El correcto midió 522,671 bytes (7054d12).

### 9) Reglas de trabajo nuevas (van al prompt)
- **Los mapas de celdas son la guía.** Antes de cada módulo Claude lee su mapa y muestra las celdas como evidencia; si
  el mapa y el Excel no coinciden, se detiene y lo presenta. No se improvisa.
- El SaaS no tiene reglas de negocio: cada pantalla recoge lo que el mapa marca como digitado y llama a la función del
  microservicio. Una función que falte se construye en la carpeta de su microservicio y se anota como pendiente.
- Si un paquete reemplaza a otro ya entregado, decirlo con los dos tamaños.
- En las pruebas, comparar montos y existencias redondeando a 2 decimales.
- Para ver el SaaS en la PC: `py -m saas_demo.ensayo` (copia temporal; nada que restaurar).

### 10) Lo que sigue: SaaS en Render (plan presentado; 1 a 2.5 h)
1. **Copia de demostración en Drive:** descargar el maestro, subirlo como archivo NUEVO con **DEMO** en el nombre
   (nunca abrirlo con Google Sheets), compartirlo con la cuenta de servicio como Editor, tomar su ID. El usuario decide
   si los teléfonos de la columna TELEFONO se quedan o se borran en la copia.
2. **Servicio web nuevo en Render (Starter):** mismo repositorio y Dockerfile; comando de arranque del SaaS
   (`gunicorn --workers 1 --threads 4 --timeout 120 saas_demo.app:app`); variables `SAAS_EXCEL_FILE_ID`,
   `SAAS_CLAVE_DESPACHO`, `SAAS_CLAVE_ADMIN` y la llave de Google como Secret File; health check `/salud`.
   El ID del maestro y las claves de Twilio NO se ponen en ese servicio.
3. **Prueba desde el navegador** con los dos usuarios (comanda, entrega, anulación en efectivo, compra) y medir tiempos.

### 11) Estimación del SaaS completo (presentada al usuario)
Base + Grupo 1: ~3 h (hecho). Render: 1 a 2.5 h. Grupo 2: 1.5 h. Grupo 3: 2 h. Grupo 4: 1.5 h. Grupo 5: 2 a 2.5 h.
Registro de eventos: 20 min. Total 11 a 13 h. Cada cierre de chat (bitácora, mapa, prompt): 30 a 40 min.

### 12) Libro de mapas V20 (LIBRO_MAPA_DE_CELDAS_PFV20.xlsx, **272583** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- Construido sobre el V19 que el usuario adjuntó al inicio (266,018 bytes; confirmó que era su última versión).
- **Regla #39 nueva** (SaaS de demostración: un módulo por microservicio en el orden del Excel, sobre la copia DEMO, tres
  candados, dos usuarios, sin reglas de negocio propias).
- Decisiones Pendientes: **#54** con "|| V20" (resuelto); bloque "ACTUALIZACIÓN V20" con **#62** (resuelto), **#63 a #69**
  (pendientes), **#70** (resuelto), **#71** (informativo) y **#72** (lección).
- Bloques "ACTUALIZACIÓN V20" en los mapas 'Ventas despacho' (comanda #70 y quién anula), 'Interface digital' (B5 con
  máximo 10 %, entregar y anular desde el SaaS) y 'Mapa_Ingreso Inventario' (pantalla de compras; producción por hornada).
- Pestaña "Cambios V20" (8 filas).
- Verificado contra V19 celda por celda: 7,304 celdas iguales, 1 con texto agregado "|| V20", 50 nuevas + pestaña nueva
  (56 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos. 53 pestañas.
- Pendiente del usuario en el libro: anotar el ID de la carpeta RESPALDO_FUENTES, la clave nueva de Maps y la de Gemini
  (#21). Cuando exista el servicio del SaaS en Render: agregar el bloque correspondiente en la Regla #37 (mapa V21).

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_actual.zip` (522,671; código 7054d12), `paquete_saas.zip` (27,850), `paquete_grupo1.zip`
  (37,793; reemplazó al de 33,289), `probar_saas_demo.py` (27,413), `env_example_ANTES_saas.txt` (1748),
  `saas_demo_ANTES_grupo1.zip` (carpeta antes del Grupo 1), `probar_saas_demo_ANTES_redondeo.py`.
- Drive: sin cambios en este chat. `.env` de la PC: sin cambios (las variables SAAS_ aún no se agregan).

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V18 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V18.md`).
2. **SaaS en Render** (punto 10): es lo primero del próximo chat.
3. **#57** (hora real de la corrida "Scheduled" del 4-oct, cerca de las 6:17 a. m.) y **#58** (evento "Desde RESPALDO manual").
4. **Grupo 2** del SaaS (hornadas, control de inventario, #63), luego Grupos 3 (#64, #65), 4 y 5 (#66 a #69) y el
   Registro de eventos. Cada grupo queda probado y cerrado antes del siguiente.
5. Después: Sesión 3 (Power BI básico y pendientes menores #51, #52, #36, #4, #17, #19, #22, #21) y Sesión 4 (entrega
   final), igual que en la V17.
6. Del usuario: renovar los dos respaldos de RESPALDO_FUENTES antes del **13-oct-2026**; elegir las dos claves del SaaS
   (van solo en el `.env` y en Render).

### Tips aprendidos hoy
- Un ZIP hecho con `git archive` guarda adentro el código del commit: sirve para comprobar que se adjuntó el correcto.
- Flask es la librería que sirve páginas web; un "blueprint" es un archivo con las pantallas de un grupo.
- Las computadoras no guardan exactos algunos decimales (8.55 - 1 no da 7.55 justo): por eso todo el proyecto redondea
  a 2 decimales.
- "Agregar a la comanda" no escribe en el Excel: la comanda vive en la sesión del navegador hasta "Registrar venta".
- El límite de un chat depende de cuánto contenido carga (código y capturas), no de las horas.

---

# CONTENIDO V17 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V17

*Última actualización: 3-oct-2026 (tarde, ~4:00 p. m. CR). Reemplaza a la V16. La V16 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V16", que a su vez trae la V15 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V17".
Subir este archivo y el libro de mapas V19 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: el chat fue el 3-oct-2026 (~10:10 a. m. a ~4:00 p. m., hora CR). La lista de corridas de GitHub Actions
muestra la hora del navegador (hora CR); el horario del archivo de la tarea va en UTC (6:17 a. m. CR = 12:17 UTC).
Las horas de este bloque son de CR y aproximadas cuando llevan "~".*

---

## ACTUALIZACIÓN V17 (3-oct-2026, ~6 h) — Hallazgo: Solera y Musi bloquean a la nube de forma intermitente. Respaldo manual en Drive (#56, Regla #38), tarea diaria a las 6:17 a. m. y guía del administrador

### Estado general
- **El sistema sigue completo en la nube**, y ahora tolera que dos fuentes externas no respondan:
  1. **Lectura directa** del sitio (un solo intento por fuente).
  2. Si el sitio no responde: **respaldo manual** en la carpeta de Drive `RESPALDO_FUENTES` (lo renueva la
     administración cada 10 días con Ctrl + S desde un navegador).
  3. Si tampoco hay respaldo válido: la celda no se toca (Regla #19).
  4. **Gemini** solo interviene para leer una **imagen nueva** de Musi que no esté en el cache (la abre desde Google).
- **Regla #38 nueva** y **Decisiones Pendientes #56 a #61** (mapa V19). #56 RESUELTO.
- **Tarea diaria movida a las 6:17 a. m. CR** (la automática del 3-oct salió a las 12:18 p. m., 5 h 18 min tarde).
- El Excel maestro **no cambió de estructura** (sin migración). 5 de 5 grupos completos, sin cambios en el modelo
  financiero.
- **La Sesión 2 (SaaS de demostración) NO se empezó**: el chat se dedicó al hallazgo. Queda para el próximo chat.

### Commits de este chat (todos con push)
- **f9b7535** — bitácora V16 (hecho al cierre del chat anterior; Render lo desplegó: servicio Live con f9b7535).
- **005bf30** — "Respaldo manual de Solera y Musi en Drive (#56) + horario 6:17 CR": 6 files changed, 837 insertions,
  33 deletions. Archivos (tamaños en la PC):
  `microservicios/grupo3_integraciones/respaldo_fuentes.py` (**12301**, nuevo),
  `probar_respaldo_fuentes.py` (**15728**, nuevo), `precios_proveedores.py` (**16853**; antes 12811),
  `competencia.py` (**41875**; antes 34775), `.github/workflows/tarea_diaria.yml` (**4603**; antes 3670),
  `.env.example` (**1748**; antes 1359). Push: `f9b7535..005bf30  main -> main`.
- **Último commit de código: 005bf30.** Pendiente al cierre: commit de esta bitácora V17 (solo este archivo).
- **Drive:** Excel maestro en **versión 163** (250,655 bytes) a las 12:19 p. m. (paso 0 de la corrida #6); mayor después
  de las corridas #6 y #7.

### 1) Revisión del Excel de Drive (adjuntado ~10:12 a. m.; 250496 bytes, 25 hojas, 0 errores)
- Última escritura: 2-oct 10:55:35 p. m. (EVT-000031). No traía nada del 3-oct.
- **Pedido de las 8:13 p. m. del 2-oct confirmado: WA-122** — 3 Baguette, retiro, SINPE Móvil, **₡2,528.07**;
  'Interface digital' fila 29, 'Detalle comanda digital' fila 41, EVT-000026; CONFIRMADO, sin entregar.
- 'Flujo caja diario' fila 219 (2-oct): ventas digitales **19,760.03** = WA-119 + WA-120 + WA-121 + WA-122; envíos 0;
  saldo acumulado **825,048.00**.
- NODO INTEGRAL C21 seguía en "Sin promocion vigente (26-09-2026)" aunque competencia corrió dos veces el 2-oct: esa
  fue la pista del hallazgo.

### 2) Hallazgo — en GitHub Actions "verde" no significa "leyó todo"
- 'RegistroEventos' filas 31 a 35 (corridas del 2-oct desde GitHub Actions), todas con estado **PARCIAL**:
  - PRECIOS_PROVEEDORES (EVT-000027, 028, 030): "Sin cambio: 5 | Fallidos 11 (la respuesta no es JSON)". Los 11
    fallidos son los productos de **Solera**; los 5 de MaxiPalí salieron bien.
  - COMPETENCIA (EVT-000029, 031): "Musi: FALLO — ninguna consulta respondio".
- Desde la PC los mismos comandos habían dado OK (EVT-000006, 007 y 012).
- GitHub pone el check verde porque el programa terminó sin caerse: los microservicios anotan el fallo como PARCIAL y
  terminan normalmente (así se diseñaron, Regla #19). **No hubo daño**: las celdas conservaron su último valor.
- Corrida **#5** (manual, 3-oct 10:40 a. m.; 4 min 31 s): precios 3 min 28 s = 5 de 16 (L01 1,050; Q01 1,800; Ma01 1,000;
  CC01 750; Na01 920; los 11 de Solera "la respuesta no es JSON"); competencia 11 s = cache de Drive con 15 imágenes,
  "Musi FALLO de conexion: ninguna consulta respondio"; NODO INTEGRAL 10 s.

### 3) Diagnóstico de la causa (pestaña Shell de Render, ~11:30 a. m.; solo consulta, sin tocar el Excel)
- `precios_proveedores --solo-consultar` desde Render: **mismo resultado que GitHub**, 5 de 16. No es un problema de
  GitHub ni del código.
- Qué responde cada sitio a un servidor de nube:
  - **Solera:** `202 text/html nginx` y una página que redirige a `/.well-known/sgcaptcha/` (parámetro `y=ipr`).
  - **Musi:** `202 text/html cloudflare` y la misma redirección a `/.well-known/sgcaptcha/`.
- Es la **verificación contra robots** del servicio donde están alojados los dos sitios, y decide por la **reputación de
  la dirección** desde donde llega la visita ("ipr"). Una conexión de casa pasa; un servidor de nube puede no pasar.
- Descartado: cambiar la forma de la petición (el filtro mira la dirección), mudarse a otra nube, y hacer que el robot
  resuelva o salte la verificación (**Claude no lo construye**: es la protección del sitio del proveedor y no se puede
  defender en una tesis). Tampoco scraping: la protección cubre todo el sitio.

### 4) Prueba: ¿puede leer Google (Gemini) lo que la nube no puede? — idea del usuario
- Archivo de prueba `probar_gemini_url.py` (**8425** bytes; se corrió desde Escritorio\download, **no está en el
  repositorio**). Usa la herramienta oficial **URL context** de Gemini: la visita al sitio la hace Google.
- 1.ª corrida (~11:50 a. m.): prueba 1 **PASA** (`HARINA FLORES 25 KG FUERTE | 0016 | 17510 | 0`, igual al Excel; el SKU
  y el precio no iban en la pregunta). Pruebas 2 a 4: **HTTP 503** ("alta demanda") en los tres modelos
  (gemini-3.8-flash, 3.7-flash y 3.6-flash).
- 2.ª corrida (~12:08 p. m.), gemini-3.8-flash, las cuatro lecturas con `URL_RETRIEVAL_STATUS_SUCCESS`:
  Solera catálogo PASA; Solera página principal PASA; **Musi catálogo de páginas PASA** (modified 2025-06-23T23:10:02,
  9 imágenes, primera `.../2024/09/panaderia-reposteria-banner-2.jpg`); Musi página pública NO PASA (Google recibió la
  página, pero sin promociones: las arma el navegador). Resumen: Solera 2 de 2, Musi 1 de 2.
- **Conclusión: la conexión por Google sirve para las dos fuentes.** Queda como evidencia y como trabajo futuro (#60).
- Gemini de pago (pregunta del usuario, con búsqueda en la documentación y el foro oficial): el pago sube el límite de
  cupo (error 429), pero **no quita la saturación** (error 503), que depende de la capacidad de Google. Google recomienda
  reintentar con pausas crecientes. **Decisión del usuario: no pagar Gemini.**

### 5) Decisiones del usuario (en orden)
1. "Necesito que funcione en la nube y no en mi PC."
2. Probar la lectura por Gemini (sí) y, como otra idea suya, un **respaldo escrito** con la información web de Solera y
   Musi para cuando falle la conexión.
3. Lo **funcional** será el respaldo manual (guardar desde el navegador con Ctrl + S); la prueba de Gemini queda como
   evidencia de que la conexión sirve.
4. Decisiones A a D, "todo como recomiendas": **A** carpeta nueva en Drive; **B** un respaldo de más de 10 días se usa
   igual, pero queda marcado; **C** una imagen nueva de Musi la lee Gemini por Google; **D** horario a las 6:17 a. m.
5. El usuario pensaba en otro orden (Gemini primero y luego el respaldo). Se aclaró y decidió: **"un alto: es mejor usar
   lo que está ya listo y probado"**. Al final: "como está hecho ya mismo está perfecto y funcional".
6. La administración actualiza los dos archivos cada 10 días. Cerrar con bitácora V17 (opción A) y dejar la Sesión 2 para
   un chat nuevo.

### 6) Lo construido — `paquete_56.zip` (**38237** bytes)
- **`respaldo_fuentes.py` (nuevo):** busca en la carpeta el archivo válido más nuevo de cada fuente
  (`respaldo_solera...`, `respaldo_musi...`). **Solo lee.** Drive: carpeta `RESPALDO_FUENTES_FOLDER_ID` (acepta el ID o
  el enlace); con `--local`: `datos_prueba/respaldo_fuentes/`. Descarta un archivo que no sea el catálogo (por ejemplo,
  la página de verificación guardada por error) y usa el anterior válido. Más de 10 días = "RESPALDO VIEJO".
  Comando de revisión: `py -m microservicios.grupo3_integraciones.respaldo_fuentes --ver` (con o sin `--local`).
- **`precios_proveedores.py`:** si Solera no responde (sin conexión, error del servidor o verificación), toma el precio
  del respaldo **por SKU**. Tras el primer fallo ya no le insiste al sitio en esa corrida (antes eran 4 intentos por
  producto: de ahí los 3 min 28 s). El respaldo solo cubre fallos de **conexión**: un producto que ya no existe en el
  catálogo FALLA y no se tapa. El evento agrega "Desde RESPALDO manual N: archivo del dd-mm-aaaa (N dias)"; con respaldo
  viejo el estado es PARCIAL.
- **`competencia.py`:** (a) si la página de Musi no responde, usa la página del respaldo; el texto que se escribe en
  NODO INTEGRAL C21 y Tabla Combos D11 lleva la **fecha del dato** (la del respaldo); (b) si el robot no puede descargar
  una imagen nueva, se la pide a Gemini por su dirección (URL context); si Google tampoco puede, lectura incompleta;
  (c) modo nuevo `--probar-vision-url <URL>`; (d) el mensaje de fallo ahora incluye el motivo de cada página.
- **`tarea_diaria.yml`:** horario `17 12 * * *` (6:17 a. m. CR) y la variable `RESPALDO_FUENTES_FOLDER_ID` (6.º Secret).
  El nombre de la tarea sigue siendo "Tarea diaria 7:00 CR" para no partir el historial.
- **`.env.example`:** documenta la variable nueva.
- `datos_prueba/respaldo_fuentes/`: las dos muestras del usuario (no van a GitHub).
- Sin carpeta de respaldos configurada, todo se comporta como antes.

### 7) Pruebas
- **Muestras reales del usuario (Ctrl + S desde su navegador):**
  - `respaldo_solera.json` (**33515** bytes): 11 de 11 productos, todos con el mismo precio que el Excel (Me01 1,815;
    Ha01 17,510; Az01 16,490; Sa01 3,310; Ma02 28,845; Le01 735; Me02 4,140; Bo01 2,400; Bu01 7,275; Ca01 20,755;
    Ma03 15,390) y en existencia. Dirección: catálogo de Solera con los 11 SKU en una sola consulta
    (`?sku=0549,0016,0021,0995,0030,0137,0957,0516,0179,0474,0034&per_page=20`).
  - `respaldo_musi.json` (**16155** bytes): página 8497, modified 2025-06-23T23:10:02, **9 imágenes, las 9 visibles**
    (1 banner de 2024/09 y 8 de 2025/06).
  - `cache_competencia.json`: 15 lecturas (9 de Musi, 6 de Musmanni). Las 9 imágenes del respaldo están en el cache.
- **En el entorno de Claude (sin internet):** prueba local 23 de 23 OK; Drive simulado 11 de 11 OK; corrida completa
  sobre una copia del Excel maestro con LibreOffice: 16 de 16 precios (11 desde el respaldo), evento OK, Musi "Sin
  promocion vigente (03-10-2026)", 0 errores de fórmula.
- **En la PC del usuario:**
  - `py -m microservicios.grupo3_integraciones.probar_respaldo_fuentes` → **RESULTADO: TODO OK (0 lineas ERROR)**.
  - `respaldo_fuentes --ver` (carpeta real de Drive): solera `respaldo_solera.json del 03-10-2026 (0 dias) -> 11
    producto(s)`; musi `respaldo_musi.json del 03-10-2026 (0 dias) -> 1 pagina(s)`; "LISTO".
  - `competencia --probar-vision-url .../2025/06/DC-MAYO-JUNIO-25-05-1024x1024.jpg` (~2:34 p. m.): gemini-3.8-flash, un
    aviso de saturación (503) y reintento a los 20 s; leyó **"Rollo de Canela Break", 2500, "*Válido por tiempo
    limitado."**, igual al cache; decisión VENCIDA. No tocó el Excel ni el cache.

### 8) Puesta en la nube
- **Drive:** carpeta `RESPALDO_FUENTES` en "Mi unidad", compartida con la cuenta de servicio como **Lector**, con los dos
  archivos (respaldo_musi.json 16 KB, respaldo_solera.json 33 KB). El ID de la carpeta vive en el `.env`
  (`RESPALDO_FUENTES_FOLDER_ID`, con el enlace completo) y en GitHub.
- **GitHub:** 6.º Secret `RESPALDO_FUENTES_FOLDER_ID` (ahora son 6: CACHE_COMPETENCIA_FILE_ID, EXCEL_MAESTRO_FILE_ID,
  GEMINI_API_KEY, GOOGLE_SERVICE_ACCOUNT_JSON, RESPALDO_FUENTES_FOLDER_ID, RESPALDO_MAESTRO_FILE_ID).
- **Corrida #7** (manual, ~3:00 p. m., código 005bf30): verde, **2 min 33 s**. Precios 1 min 21 s: **16 de 16**, los 11
  de Solera "via slug" (lectura directa), "Sin cambio: 16". Competencia 12 s: Musi leído directo, 9 imágenes del cache,
  "Sin promocion vigente (03-10-2026)", sin cambios en las celdas. NODO INTEGRAL 11 s.
- **En la #7 el respaldo no hizo falta** (los dos sitios contestaron), así que **todavía no se ha visto el respaldo
  trabajando en la nube** (#58). Sí está probado en la PC con la carpeta real de Drive.

### 9) El horario — la corrida automática salió 5 h 18 min tarde
- A las ~10:36 a. m. la lista de GitHub Actions mostraba solo las 4 corridas manuales del 2-oct: la automática de las
  7:00 no había salido.
- **Corrida #6, "Scheduled", 12:18 p. m.** (programada 7:00 a. m.; **5 h 18 min de atraso**), 4 min 59 s, código f9b7535
  (el de antes del respaldo): respaldo diario OK (maestro versión 163, 250,655 bytes, copiado 12:19, huella igual);
  NODO: clima 'Despejado' → 'Lluvia', temperatura 30 → 30.1, combos activos co02 y co03; precios 5 de 16 (**Solera
  bloqueada**); competencia: **Musi sí respondió** y cambió C21 y D11 de "Sin promocion vigente (26-09-2026)" a
  "Sin promocion vigente (03-10-2026)".
- GitHub advierte en su documentación que las tareas programadas pueden atrasarse con mucha carga, sobre todo al inicio
  de cada hora. Por eso se movió a un minuto fuera de la hora en punto: **6:17 a. m. CR**.
- Pendiente (#57): comprobar el 4-oct que aparece una corrida "Scheduled" cerca de las 6:17. Si vuelve a atrasarse por
  horas: opción 1C (que otro servicio dispare la tarea; requiere una clave nueva de GitHub; decisión del usuario).

### 10) El bloqueo es intermitente y cada sitio decide por su cuenta (dato para la tesis)

| Corrida | Hora (3-oct) | Código | Solera | Musi | Duración |
|---|---|---|---|---|---|
| #5 manual | 10:40 a. m. | f9b7535 | bloqueada | bloqueado | 4 min 31 s |
| Render (Shell) | ~11:30 a. m. | f9b7535 | bloqueada | bloqueado | — |
| #6 automática | 12:18 p. m. | f9b7535 | bloqueada | pasó | 4 min 59 s |
| #7 manual | ~3:00 p. m. | 005bf30 | pasó | pasó | 2 min 33 s |

- Las corridas del 2-oct en la noche (#3 y #4) también quedaron con los dos sitios bloqueados (eventos PARCIAL).
- No se puede saber de antemano si un sitio va a contestar: por eso el robot hace **un intento directo** (si pasa, el
  dato es del día) y, si no, usa el respaldo. En una corrida como la #6, con el código nuevo, los 11 precios de Solera
  habrían salido del respaldo en vez de quedar en fallo.
- Frase para el reporte: una automatización real depende de los permisos de terceros; aquí está demostrado con evidencia
  (capturas de GitHub Actions, del Shell de Render y de la Terminal).

### 11) Mediciones (dato para la tesis)
- Tarea diaria con Solera bloqueada: 4 min 23 s a 4 min 59 s (precios ~3 min 28 s por los reintentos).
- Tarea diaria con el código nuevo y los sitios respondiendo: **2 min 33 s** (precios 1 min 21 s; competencia 12 s;
  NODO 11 s).
- Atraso de la tarea programada de GitHub el 3-oct: **5 h 18 min** (7:00 → 12:18).
- Gemini (nivel gratuito) el 3-oct: 503 en los tres modelos flash a las ~11:50 a. m.; respuesta normal a las ~12:08 p. m.
  y a las ~2:34 p. m. (con un reintento de 20 s).

### 12) Procedimiento del administrador (cada 10 días, 5 a 10 minutos, solo navegador)
1. Abrir la dirección del catálogo de Solera (los 11 SKU), Ctrl + S, nombre que empiece por `respaldo_solera`.
2. Abrir la dirección de la página de promociones de Musi, Ctrl + S, nombre que empiece por `respaldo_musi`.
3. Arrastrar los dos archivos a la carpeta `RESPALDO_FUENTES` de Drive.
- Las dos direcciones exactas están en el encabezado de `respaldo_fuentes.py` y en la guía.
- Reglas: el nombre es lo que cuenta; no hace falta borrar los anteriores; no abrirlos con aplicaciones de Google.
- Si se olvida: el robot usa el último respaldo y el evento queda PARCIAL con "RESPALDO VIEJO".
- No es tarea del administrador: un insumo nuevo de Solera o un cambio de SKU (hay que ajustar la lista de SKU en el
  robot y en la dirección).
- **Guía de una página entregada:** `Guia_administrador_respaldo_Solera_Musi.docx` (**11278** bytes).

### 13) Correcciones a la V16 (el texto de la V16 queda abajo sin cambios)
- **|| V17:** la V16 dice "#14 resuelto" y "el sistema funciona completo en la nube". Precisión: la tarea diaria corre
  sola, pero el 2-oct y el 3-oct en la mañana **no leyó a Solera ni a Musi** (eventos PARCIAL). Queda resuelto con el
  respaldo manual (#56) y el horario nuevo; #57 y #58 siguen en observación.
- **|| V17:** el "pedido de las 8:13 p. m., número y total por confirmar" es **WA-122, ₡2,528.07** (punto 1).
- **|| V17:** "tarea diaria a las 7:00 a. m. CR" pasa a **6:17 a. m. CR**; "5 Secrets" pasa a **6**.

### 14) Reglas de trabajo nuevas (van al prompt)
- En GitHub Actions el check verde solo dice que el programa terminó: el resultado real se lee en 'RegistroEventos'
  (OK / PARCIAL) o abriendo el paso.
- **Cada recuadro lleva arriba su destino** (Terminal, barra del navegador, campo de una página o archivo `.env`). Lo que
  no es para la Terminal no se entrega sin decir dónde va: el usuario pegó en la Terminal una dirección, un correo y un
  renglón del `.env` (errores en rojo, sin daño).
- Cuando se mezclan navegador, `.env` y Terminal: **un destino por mensaje**.
- Para abrir un paso de una corrida: decir el nombre exacto del paso y su posición en la lista.
- Supuestos de Claude que resultaron equivocados y se corrigieron en el chat: (a) "la corrida automática se descartó"
  (salió 5 h 18 min tarde); (b) "en la #6 veremos el respaldo" (la #6 corrió con el código viejo); (c) el orden que
  Claude construyó primero no era el que el usuario tenía en mente (se aclaró; el usuario decidió dejar lo construido);
  (d) "el nivel de pago de Gemini baja bastante la saturación" (la evidencia dice que no la quita).

### 15) Libro de mapas V19 (LIBRO_MAPA_DE_CELDAS_PFV19.xlsx, **266018** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- **Regla #38 nueva** (fuentes que bloquean a la nube: lectura directa y, si falla, respaldo manual en Drive). Reglas
  #18, #19 y #37 con texto agregado "|| V19". Decisiones Pendientes: #14 corregido y #53 confirmado con "|| V19"; bloque
  "ACTUALIZACIÓN V19" con **#56** (resuelto), **#57** y **#58** (pendientes), **#59** (informativo: evidencia), **#60**
  (trabajo futuro: Gemini lee Solera cada día) y **#61** (lección). Bloques "ACTUALIZACIÓN V19" en los mapas
  'NODO INTEGRAL', 'Tabla combos dinamicos' y 'Mapa TABLA CRUCE MAESTRA MP'. Pestaña "Cambios V19" (8 filas).
- Verificado contra V18 celda por celda: 7,214 celdas iguales, 5 con texto agregado "|| V19", 30 nuevas + pestaña nueva
  (56 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos. 52 pestañas.
- Pendiente del usuario en el libro: anotar el ID de la carpeta RESPALDO_FUENTES, la clave nueva de Maps y la de Gemini
  (#21).

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_8cccf47.zip`, `repo_actual.zip` (código f9b7535), `probar_gemini_url.py` (8425),
  `respaldo_solera.json`, `respaldo_musi.json`, `paquete_56.zip` (**38237**), `competencia_ANTES_56.py` (34775),
  `precios_proveedores_ANTES_56.py` (12811), `env_example_ANTES_56.txt` (1359),
  `Guia_administrador_respaldo_Solera_Musi.docx`. El respaldo de `tarea_diaria.yml` no se hizo (se repitió el de
  `.env.example`); la versión anterior está en Git (f9b7535).
- datos_prueba: `respaldo_fuentes\respaldo_solera.json` y `respaldo_musi.json` (muestras de la prueba local).
- Drive: carpeta nueva `RESPALDO_FUENTES` con los dos respaldos; el Excel maestro, `CACHE_COMPETENCIA.xlsx` y
  `RESPALDO_DIARIO_MAESTRO.xlsx` siguen igual.
- `.env` de la PC: renglón nuevo `RESPALDO_FUENTES_FOLDER_ID=` (renglón 38).

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V17 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V17.md`).
2. **#57:** comprobar en GitHub → Actions que hay una corrida "Scheduled" cerca de las 6:17 a. m. (anotar la hora real).
3. **#58:** revisar 'RegistroEventos' unos días hasta ver un evento con "Desde RESPALDO manual" (el respaldo trabajando en
   la nube).
4. **Sesión 2 — SaaS de demostración** (copia del Excel, usuario y contraseña, sin tocar el maestro) con botón
   "Entregado" que llama `marcar_entregado` y el **endpoint de anulación** solo para administración (#54). Claude presenta
   primero las opciones con su costo en tiempo.
5. **Sesión 3 — Power BI básico** y pendientes menores: #51, #52, #36, #4, #17, #19, #22, #21. WA-121 y WA-122 siguen sin
   entregar.
6. **Sesión 4 — Entrega final** (sin video, sin ZIP): revisión de seguridad del historial de commits, README y guía del
   repositorio para los tutores, invitarlos como colaboradores de solo lectura, instrucciones del Sandbox (comprobar que
   desde México pueden). Trabajo escrito: normativa, #26, tiempos medidos (PC, Render gratis y Starter, GitHub Actions),
   tasa de abandono, **el hallazgo del bloqueo intermitente y el respaldo manual**, camino a producción (número propio,
   autorización de Solera para el servidor del robot, Gemini leyendo las fuentes cada día (#60), depuración del cache,
   rotación de claves, atención humana).
7. Del usuario: renovar los dos respaldos antes del **13-oct-2026**; anotar en el libro de mapas el ID de la carpeta, la
   clave de Maps y la de Gemini; decidir si borra `env_ANTES_43.txt`.

### Tips aprendidos hoy
- Ctrl + S sobre una dirección que muestra texto con llaves y comillas guarda ese mismo texto como archivo: es el dato del
  sitio, sin que nadie lo digite.
- La dirección de una carpeta de Drive no es una clave: solo la abren las cuentas con las que se compartió.
- En Git: rojo = cambio que aún no se eligió; `??` = archivo nuevo que Git no conoce; verde = elegido para el commit
  (M modificado, A agregado). `git add` mete en el sobre solo lo que se nombra; por eso nunca "git add .".
- En Render, la pestaña **Shell** (menú MANAGE del servicio) es una Terminal dentro del servidor; el buscador de Deploys
  no ejecuta nada.
- En GitHub Actions, la corrida automática dice **Scheduled**; las de botón dicen "Manually run".
- Los avisos de LF y CRLF de `git add` en Windows no son errores.

---

# CONTENIDO V16 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V16

*Última actualización: 2-oct-2026 (noche, ~11:20 p. m. CR). Reemplaza a la V15. La V15 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V15", que a su vez trae la V14 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V16".
Subir este archivo y el libro de mapas V18 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: el chat fue el 2-oct-2026 (~3:40 p. m. a ~11:20 p. m., hora CR). Render y el bot muestran hora de CR;
GitHub Actions muestra UTC (7:00 a. m. CR = 13:00 UTC). Los archivos de los paquetes llevan fecha 3-oct (UTC).*

---

## ACTUALIZACIÓN V16 (2-oct-2026, ~7.5 h) — Hosting completo: Render Starter, tarea diaria en GitHub Actions, cache de competencia en Drive y respaldo diario

### Estado general
- **El sistema ya funciona completo en la nube, sin la PC del usuario:**
  1. **Bot de WhatsApp** en Render, plan **Starter** (no se duerme; escrituras en 12-13 s).
  2. **Tarea diaria** en GitHub Actions, todos los días a las **7:00 a. m. CR**: respaldo diario → NODO INTEGRAL →
     precios de proveedores → competencia (Musi).
  3. **Cache de Gemini** en el Excel `CACHE_COMPETENCIA.xlsx` de Drive (15 lecturas).
  4. **Respaldo diario** del Excel maestro en `RESPALDO_DIARIO_MAESTRO.xlsx` de Drive.
- 5 de 5 grupos completos (sin cambios en el modelo financiero). Solo queda #26 del modelo (redacción del usuario).
- Cerrados hoy: **#14, #20, #42, #48, #49** y el respaldo diario (#50). **#43** quedó parcial y cerrado por decisión del
  usuario (ver punto 5).

### Commits de este chat (todos con push)
- **8cccf47** — bitácora V15 (hecho al cierre del chat anterior; confirmado con la captura del push).
- **8836590** — "Cache de competencia en Excel de Drive (#20) + tarea diaria 7:00 CR en GitHub Actions (#14)":
  `.github/workflows/tarea_diaria.yml` (3199, nuevo), `microservicios/grupo3_integraciones/cache_competencia.py`
  (**16705**, nuevo), `probar_cache_excel.py` (**6869**, nuevo), `competencia.py` (**34775**; antes 32619).
  4 files changed, 634 insertions, 3 deletions.
- **e511666** — "Respaldo diario del Excel maestro en Drive (paso 0 de la tarea diaria)":
  `conexion_drive/respaldo_diario.py` (**7148**, nuevo) y `.github/workflows/tarea_diaria.yml` (**3670**).
- **Último commit de código: e511666** (`HEAD -> main, origin/main`; `git status --short` vacío).
  Pendiente al cierre: commit de esta bitácora V16 (solo este archivo).
- **Drive:** Excel maestro en **versión 142** (250,383 bytes) a las ~10:45 p. m.; mayor después de la última corrida
  de la tarea (~10:56 p. m.).

### 1) #48 resuelto — copia local restaurada
- `datos_prueba\maestro_prueba.xlsx` restaurada desde `maestro_prueba_RESPALDO_20261001_190238.xlsx`:
  **249857** bytes, 29-sep 4:18:28 p. m. (igual a la copia sana anotada en la V14).
- La copia rota se conserva como `datos_prueba\maestro_prueba_ANTES_restaurar48.xlsx` (239971, 1-oct 7:02:39 p. m.).
- Verificación de solo lectura: NODO INTEGRAL C19 y C20 = OK (`_xlfn.TEXTJOIN`).

### 2) Revisión del Excel de Drive (descargado ~4:00 p. m., 249513 bytes, 25 hojas, 0 errores)
- **WA-119** (2-oct 12:03 a. m., 7,098.68, ENTREGADO 12:06 a. m., ESTADO CONFIRMADO) y **WA-120** (9:12 a. m.,
  1 Combo 1, 3,377.76) en 'Interface digital' filas 26-27; 'Detalle comanda digital' filas 37-39; eventos
  EVT-000021 a EVT-000023.
- 'Flujo caja diario' fila 219 (2-oct): ventas digitales **10,476.44** = 7,098.68 + 3,377.76; envíos 0; saldo acumulado
  805,287.97 → **815,764.41**. La Regla #35 funciona: el pedido entregado sigue sumando.
- 'Control de Inventario': Baguette 246 → 237 (3 Combo 1 × 3), CupCake 124 → 123, leche 8.55 → 5.55.
- NODO INTEGRAL seguía con los datos del 1-oct (por eso hacía falta la tarea diaria, #14).

### 3) #42 y #49 resueltos — Render Starter (decisión del usuario, opción C)
- Opciones presentadas: A ping con cron-job.org, B UptimeRobot, C plan pagado. **El usuario eligió C** para que el bot
  no se duerma y no repetir la pérdida de tiempo que tuvo con la prueba gratis de Twilio.
- **Starter: US$7 al mes**, 0.5 CPU, 512 MB; cobro prorrateado por tiempo y al final del mes (no prepago); se puede
  volver a Free cuando se quiera. Tarjeta registrada en Billing; el **workspace sigue en Hobby** (gratis); Autoscaling Off.
- Activo desde el **2-oct-2026 6:04:44 p. m. CR** (deploy "Compute plan updated", 11.2 s, commit 8cccf47). La dirección
  no cambió. `/salud` responde `{"escrituras_pendientes":0,"estado":"ok"}`.
- Prueba de no-sueño: último tráfico 6:11 p. m.; "Hola" a las 6:39 p. m. (**28 min después**) y el bot contestó el menú
  en el mismo minuto.
- No tocar: botón "Upgrade" de la barra de arriba ni "Update Plan" de Billing (son planes del workspace, por usuario).

### 4) Mediciones (dato para la tesis; horas exactas de los Logs de Render)
| Escritura | PC | Render gratis (0.1 CPU) | Render Starter (0.5 CPU) |
|---|---|---|---|
| Registrar pedido | 7.9 s | < 1 min (WA-119) | **12.2 s** (WA-121) y **12.3 s** (pedido de las 8:13 p. m.) |
| Marcar entregado | 7.9 s | < 1 min (WA-119) | **13.3 s** (WA-120) |
| Arranque en frío | — | ~1 min; se pierde el primer mensaje | no hay (no se duerme) |
- WA-120 entregado: fila 18:11:16.594 → "Entregar pedido WA-120: guardado (intento 1)" 18:11:29.342 → Twilio 201
  18:11:29.889.
- WA-121 registrado: fila 18:47:22.644 → "guardado (intento 1)" 18:47:34.729 → Twilio 201 18:47:34.885.
- Pedido de las 8:13 p. m.: fila 20:13:33.283 → guardado 20:13:45.296 → Twilio 201 20:13:45.585.
- Todas entraron al **primer intento** (el arreglo #37 de la huella md5 sigue bien).
- Tarea diaria en GitHub Actions: **4 min 20 s** (LibreOffice 24 s, librerías 5 s, NODO 6 s, precios 3 min 27 s,
  competencia 11 s); con el respaldo diario, **4 min 50 s**.
- Valoración del usuario: "es mucho más rápido el bot ahora".

### 5) #43 seguridad — hecho en parte y CERRADO POR DECISIÓN DEL USUARIO
- **Hecho:** clave nueva de Google Maps **"Maps Routes bot oct-2026"** (proyecto Robot Panaderia Canas, restringida
  solo a **Routes API**, sin restricción de aplicación), puesta a mano en el `.env` de la PC (un solo renglón
  GOOGLE_MAPS_API_KEY) y en Render (reinicio 8:09:40 p. m.). Probada: PC `google_maps 10.4306 -85.0967` = TODO BIEN,
  0.8 km, 3 min, envío 400 colones; el bot respondió 171.64 km (fuera del radio de 10 km) desde la ubicación del usuario.
- **Decisión del usuario:** la clave vieja "Maps Platform API Key" (6-ago-2026, 35 API) **NO se borra** ("así se queda")
  y el **Auth Token de Twilio NO se cambia**. Motivo: la rotación de claves no tiene sentido en un documento académico y
  le costó mucho tiempo. No queda como pendiente.
- Respaldo del `.env` anterior: `Escritorio\download\env_ANTES_43.txt` (1598 bytes). **Contiene claves**: nunca
  adjuntarlo ni copiarlo al proyecto; el usuario decide si lo borra.
- El `.env` de la PC **no tiene** variables TWILIO_ ni BOT_ (viven solo en Render).
- Pendiente del usuario (libro de mapas, Regla #15): anotar la clave nueva de Maps en NODO INTEGRAL.

### 6) #14 resuelto — tarea diaria en GitHub Actions
- Archivo `.github/workflows/tarea_diaria.yml`: `cron "0 13 * * *"` (13:00 UTC = 7:00 a. m. CR) + botón "Run workflow";
  `ubuntu-24.04`, Python 3.12, LibreOffice 24.2 (`libreoffice-calc libreoffice-core fonts-dejavu-core`), acciones
  `actions/checkout@v6` y `actions/setup-python@v6` (Node 20 ya fue retirado de los runners).
- Pasos: **0) Respaldo diario** → **1) NODO INTEGRAL** → **2) Precios de proveedores** → **3) Competencia (Musi)**.
  Son los mismos comandos de la PC (`python -m ...`); no hubo que cambiar código del robot para esto (el código ya leía
  `GOOGLE_SERVICE_ACCOUNT_JSON`). Si un paso falla, los siguientes igual se intentan y la corrida queda en rojo.
- **5 Secrets** (GitHub → Settings → Secrets and variables → Actions): EXCEL_MAESTRO_FILE_ID,
  GOOGLE_SERVICE_ACCOUNT_JSON, CACHE_COMPETENCIA_FILE_ID, GEMINI_API_KEY, RESPALDO_MAESTRO_FILE_ID. Valores fijos en el
  archivo (no secretos): ORIGEN_EXCEL=drive, RECALCULAR=si, ZONA_HORARIA, LIMITE_CAMBIO_PRECIO=0.5.
- Corridas manuales: #1 y #2 en rojo (Secrets mal pegados, ver punto 9); **la tercera en verde (4 min 20 s)** y la
  cuarta, ya con el respaldo, **en verde (4 min 50 s)**.
- Clima y tipo de cambio quedan **1 vez al día** (la opción de cada 30 min no se programó).
- Revisar en la próxima sesión que la corrida automática de las 7:00 del 3-oct salió en verde (pestaña Actions).

### 7) #20 resuelto — cache de competencia en el Excel de Drive (opción B, decisión del usuario)
- Opciones: A cache de GitHub (sin código), B el Excel `CACHE_COMPETENCIA.xlsx` de Drive (diseño del usuario del 25-sep),
  C dejar competencia fuera. El usuario confirmó que el cache **debe estar en la nube y no en su PC** y, como la
  diferencia era cerca de una hora, eligió **B**.
- `cache_competencia.py` (nuevo): lee y escribe TablaCache (14 columnas, una fila por lectura = historial acumulado;
  si una imagen se relee vale la más reciente; ID_LECTURA `LEC-000001`...). Drive: `CACHE_COMPETENCIA_FILE_ID`;
  con `--local`: `datos_prueba\CACHE_COMPETENCIA_prueba.xlsx`; **si no hay Excel de cache se sigue usando el JSON**.
  Recalcula con LibreOffice (hoja Control: filas usadas, imágenes distintas). Comandos: `--ver` y `--migrar-json`.
- `competencia.py`: 3 renglones cambiados y 39 agregados (`_cargar_cache`, `_guardar_cache(cache, nueva=...)`); cada
  lectura nueva de Gemini se guarda en seguida como una fila; si guardar falla, avisa y reintenta al final.
- Pruebas: aquí con LibreOffice 24.2 y en la PC: `probar_cache_excel` **TODO OK (0 lineas ERROR)**; `--migrar-json --local`
  15 lecturas y 2a corrida "Ya estaba migrado"; `competencia --solo-consultar --local` con la página real de Musi =
  9 imágenes, las 9 desde el cache, **0 llamadas a Gemini**, "Sin promocion vigente (02-10-2026)".
- **Drive:** `--migrar-json` = de 0 a **15 lecturas de 15 imágenes** (ORIGEN "migrado JSON", LEC-000001 a LEC-000015);
  `competencia --solo-consultar` leyó el cache desde GOOGLE DRIVE con el mismo resultado. El JSON no se borró.
- **Límite de filas (hoja Control B4 = 1,000,000): decisión del usuario = no se construye nada.** El cache se creó para
  que Gemini no se atrase; con 7-9 imágenes no tiene sentido en un documento académico; en producción la empresa lo
  depuraría. Aclaración del usuario: la regla "nada se borra" aplica a los archivos intermedios mientras se llega al
  correcto; no obliga a conservar para siempre datos viejos.
- Para la tesis: lo que usa el negocio (promoción vigente de Musi y evento diario COMPETENCIA) queda en el Excel maestro;
  el detalle de cada imagen leída queda en el cache.

### 8) Respaldo diario del Excel maestro (#50) — opción A, decisión del usuario
- Opciones: A archivo de respaldo en Drive, B copia descargable en GitHub (los tutores podrían bajar el maestro).
  **El usuario eligió A.**
- El usuario creó en Drive **`RESPALDO_DIARIO_MAESTRO.xlsx`** (Hacer una copia del maestro, misma carpeta, compartido
  como Editor con la cuenta de servicio). Su ID va en el Secret RESPALDO_MAESTRO_FILE_ID (y en el libro de mapas).
  **Nunca abrirlo con Google Sheets** (su enlace lo abriría así).
- `conexion_drive/respaldo_diario.py`: copia el maestro **byte por byte** (solo lo lee) al archivo de respaldo; acepta
  el ID o el enlace. Seguros: el destino no puede ser el maestro, su nombre debe contener **RESPALDO**, debe ser .xlsx
  y editable; después de subir compara la huella md5; si ya es idéntico no sube nada. Como se escribe una vez al día, el
  historial de versiones de Drive de ese archivo queda con una versión por día.
- Pruebas: aquí con un Drive simulado, 9 de 9 casos; en la PC `py -m conexion_drive.respaldo_diario <ID>` = maestro
  versión 142, 250,383 bytes, "El respaldo ya era identico al maestro: no se subio nada"; corrida de GitHub en verde.
- Para recuperar un día: Drive → clic derecho en el respaldo → Información del archivo → Administrar versiones.

### 9) Problemas encontrados y resueltos (en orden)
1. Al crear la clave de Maps el usuario entró primero a **Google AI Studio** (claves de Gemini) en vez de Google Cloud
   Console: se canceló sin crear nada. La página correcta es APIs y servicios → Credenciales del proyecto.
2. En el `.env` la clave nueva se agregó como un **segundo renglón** con el mismo nombre (quedaron la vieja y la nueva):
   se borró el renglón viejo. El nombre de la variable no cambia; lo que cambia es el valor.
3. Varias capturas del `.env` abierto mostraron el **inicio** de claves cortadas por el borde (ninguna completa).
4. Los comandos largos de PowerShell para editar el `.env` confundieron; el usuario prefirió editar a mano.
5. **Secrets de GitHub con su propio nombre como valor** (el portapapeles se pisó al copiar el nombre desde el chat):
   corrida #1 `JSONDecodeError: Expecting value` (GOOGLE_SERVICE_ACCOUNT_JSON) y corrida #2 `404 File not found`
   (CACHE_COMPETENCIA_FILE_ID). Se corrigieron con el lápiz (Update secret). Pista: GitHub tapa con `***` cualquier
   texto igual al valor de un Secret.
6. Una dirección pegada con ": " delante se fue a una búsqueda de Google: las direcciones se entregan solas, en recuadro.

### 10) Pedidos de prueba en Drive (bot en Render)
- **WA-120**: 2-oct 9:12 a. m., 1 Combo 1, retiro, SINPE Móvil, ₡3,377.76; **ENTREGADO 2-oct 6:11 p. m.** (ya no se anula).
- **WA-121**: 2-oct 6:47 p. m., 2 Combo 1, retiro, SINPE Móvil, ₡6,755.52; CONFIRMADO, **sin entregar**.
- Pedido de las **8:13 p. m.**: 3 Baguette, retiro, SINPE Móvil (antes el bot rechazó el envío: 171.64 km). Número y total
  **por confirmar** (probablemente WA-122; 3 × 842.69 = 2,528.07); sin entregar.
- El Excel descargado en la tarde ya quedó viejo: para revisar el flujo, descargar de nuevo de Drive.

### 11) Hallazgos anotados (mapa V18)
- **#51 (pendiente):** el usuario percibió ~15 s en el saludo de las 6:39 p. m. (primer saludo tras el reinicio); a las
  6:44 todo respondió en el mismo minuto. El saludo lee el Excel de Drive en vivo y Twilio espera ~15 s. Los pasos de la
  conversación **no dejan renglón en los Logs** (solo las escrituras). Propuesta: registrar un renglón por mensaje con su
  duración (sin teléfono ni texto) y evaluar un cache corto del catálogo.
- **#52 (menor):** la librería de Twilio escribe ~20 renglones por mensaje enviado en los Logs (sin el token).
- **#36:** el menú del bot muestra "Margarina Numar caja de 4 Uds - 500 g" a ₡262.50 (precio de una barra).
- **#54:** el endpoint de anulación para administración se pasa a la sesión del SaaS.

### 12) Reglas de trabajo que pidió el usuario en este chat (van al prompt)
- **Claude sugiere y el usuario decide.** Cada opción se presenta con su costo en tiempo; ningún cambio de claves o de
  APIs se hace si el usuario no lo indica.
- Antes de empezar un cambio, explicar **qué es y por qué**; no avanzar "a ciegas".
- Instrucciones **breves y exactas**; preguntar la prioridad antes de tratar algo como urgente (es un proyecto académico
  con fecha de entrega y el tiempo del usuario es limitado).
- No dejar cosas a medias: lo que se empieza se cierra, o se deja cerrado por decisión del usuario.

### 13) Libro de mapas V18 (LIBRO_MAPA_DE_CELDAS_PFV18.xlsx, **259144** bytes) — FUERA DEL REPO (tiene claves, Regla #15)
- **Regla #37 nueva** (operación en la nube: bot en Render Starter, tarea diaria en GitHub Actions, cache y respaldo en
  Drive, dónde viven las claves). Decisiones Pendientes: #14, #20, #42, #48, #49 RESUELTOS; #43 cerrado por decisión del
  usuario; #36 con evidencia nueva; bloque "ACTUALIZACIÓN V18" con #50 (resuelto), #51, #52, #54 (pendientes), #53
  (informativo) y #55 (lección de los Secrets). Bloques "ACTUALIZACIÓN V18" en los mapas 'NODO INTEGRAL' y 'Tabla combos
  dinamicos'. Pestaña "Cambios V18" (10 filas).
- Verificado contra V17 celda por celda: 7,115 celdas iguales, 7 con texto agregado "|| V18", 29 nuevas + pestaña nueva
  (68 celdas), **0 borradas, 0 cambiadas**, 47 comentarios y 7 tablas intactos.
- Pendiente del usuario en el libro: clave nueva de Maps (NODO INTEGRAL) y #21 (clave de Gemini).

### Respaldos y archivos (nada se borró)
- Escritorio\download: `repo_8cccf47.zip`, `paquete_20.zip` (**33026**), `competencia_ANTES_20.py` (32619),
  `paquete_respaldo.zip` (**4571**), `tarea_diaria_ANTES_respaldo.yml` (3199), `env_ANTES_43.txt` (1598, CON CLAVES).
- datos_prueba: `maestro_prueba_ANTES_restaurar48.xlsx` (239971), `CACHE_COMPETENCIA_prueba.xlsx` (copia local del cache
  con las 15 lecturas), `cache_competencia.json` (se conserva).
- Drive (carpeta del maestro): Excel maestro, `CACHE_COMPETENCIA.xlsx`, `RESPALDO_DIARIO_MAESTRO.xlsx`.

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V16 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V16.md`).
2. Comprobar en GitHub → Actions que la corrida automática de las 7:00 a. m. salió en verde.
3. **Sesión 2 — SaaS de demostración** (copia del Excel, usuario y contraseña, sin tocar el maestro) con botón "Entregado"
   que llama `marcar_entregado` y el **endpoint de anulación** solo para administración (#54).
4. **Sesión 3 — Power BI básico** y pendientes menores: #51 (registro y tiempo del saludo), #52 (Logs de Twilio), #36
   (margarina), #4, #17, #19, #22, #21. Confirmar número y total del pedido de las 8:13 p. m.; WA-121 y ese pedido siguen
   sin entregar.
5. **Sesión 4 — Entrega final** (sin video, sin ZIP): revisión de seguridad del historial de commits, README y guía del
   repositorio para los tutores, invitarlos como colaboradores de solo lectura, instrucciones del Sandbox; comprobar que
   los tutores en México pueden unirse al Sandbox. Trabajo escrito: normativa, #26, tiempos medidos (PC, Render gratis y
   Starter, GitHub Actions), tasa de abandono, camino a producción (número propio, depuración del cache, rotación de
   claves, atención humana).
6. Del usuario: anotar la clave nueva de Maps y la de Gemini (#21) en el libro de mapas; decidir si borra
   `env_ANTES_43.txt`.

### Tips aprendidos hoy
- En el `.env` cada variable va **una sola vez**; cambiar una clave es cambiar el valor, no el nombre.
- El `.env` abierto no se captura; para comprobar se usa la Terminal (largo y primeros 4 caracteres).
- Secrets de GitHub: primero el nombre y **después** copiar el valor; o usar el lápiz, donde el nombre ya viene fijo.
  Una vez guardados no se pueden volver a ver.
- "New repository secret" crea un secreto dentro del repositorio; no crea un repositorio.
- En Render: "Upgrade now" de la franja morada o Settings → Instance Type cambian la instancia; "Upgrade" de la barra de
  arriba es otro plan.
- El enlace "Copiar enlace" de un .xlsx en Drive lo abriría con Google Sheets: usarlo solo como dato (el ID está entre
  /d/ y /edit).
- `git archive --format=zip -o archivo.zip HEAD` empaqueta solo lo que Git tiene guardado (sin `.env` ni credenciales).
- Flecha hacia arriba en la Terminal repite el último comando.

---

# CONTENIDO V15 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V15

*Última actualización: 2-oct-2026 (madrugada, ~12:30 a. m. CR). Reemplaza a la V14. La V14 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V14", que a su vez trae la V13 ... V8); lo nuevo está en este bloque "ACTUALIZACIÓN V15".
Subir este archivo y el libro de mapas V17 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: el chat fue el 1-oct-2026 (~3:00 p. m. a 12:30 a. m. del 2-oct, hora CR). Los registros de Render
muestran la hora de CR (-0600).*

---

## ACTUALIZACIÓN V15 (1-oct-2026, ~9 h) — Fase de HOSTING: bot en Render + Twilio, ciclo completo probado desde el celular

### Estado general
- **El bot de WhatsApp funciona en internet**: Render (plan gratis, Docker) + Twilio (Sandbox, cuenta Full access).
  Ciclo completo probado desde el celular del usuario: saludo con promoción y menú leído en vivo del Excel de Drive →
  pedido → **WA-119 registrado en Drive** → confirmación con número → "Entregado 119" → **recibo** al cliente.
- 5 de 5 grupos completos (sin cambios en el modelo). Solo queda #26 del modelo (redacción del usuario).

### Commits de este chat (todos con push)
- **43e2460** — `Dockerfile` (1817, versión Debian, reemplazada después), `.dockerignore` (**493**) y
  `conexion_drive/drive_excel.py` (**12668**): arreglo del **#37 opción B1** (huella md5 en vez del número de versión).
- **fe9b754** — `whatsapp_bot/app.py` (**4759**) y `whatsapp_bot/config.py` (**3059**): la firma de Twilio se compara con
  la URL pública del webhook (variable `BOT_URL_PUBLICA`) y, si falla, el registro muestra las direcciones probadas.
- **a4e5754** — `Dockerfile` (**2464**): caja **Ubuntu 24.04 + LibreOffice 24.2 + Python 3.12 (venv)**.
- **Último commit de código: a4e5754.** Pendiente al cierre: commit de esta bitácora V15 (solo este archivo).
- La bitácora V14 ya estaba en GitHub (commit fbd264c).

### 1) #37 resuelto (opción B1, aprobada por el usuario) — drive_excel.py
- Antes de subir, el robot ya no compara el NÚMERO de versión de Drive (Drive lo sube solo, segundos después de cada
  subida), sino la **huella md5 del contenido**: la de los bytes descargados contra `md5Checksum` actual en Drive. Igual =
  nadie escribió → se sube; distinta = choque real → reintenta (se mantienen 3 intentos). Modo local sin cambio.
- Probado en el entorno de Claude con un Drive simulado (versión 107→110 sin cambio de contenido: guardado intento 1;
  choque real: intento 2; local: intento 1) y contra **Drive real**: NODO INTEGRAL "guardado (intento 1)" (clima Lluvia,
  tipo de cambio 460.73, combos co01 y co04). **Evidencia nueva del #37:** con UNA sola escritura Drive pasó de 107 a 110.

### 2) Render (servicio web)
- Cuenta creada con GitHub (Gigo1965); workspace "Robot Panaderia"; acceso de la GitHub App **solo** al repo
  robot-panaderia.
- Servicio **robot-panaderia**: Docker, rama main, región **Virginia (US East)**, plan **Free** (512 MB, 0.1 CPU),
  URL **https://robot-panaderia.onrender.com**, health check **/salud**, **Auto-Deploy: On Commit**.
- Variables de entorno (NOMBRES; los valores viven solo en Render y en el .env): EXCEL_MAESTRO_FILE_ID,
  GOOGLE_MAPS_API_KEY, BOT_NUMEROS_AUTORIZADOS (solo el celular del usuario por ahora), GOOGLE_SERVICE_ACCOUNT_FILE
  (= /etc/secrets/clave_google.json), TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_WHATSAPP_FROM
  (= whatsapp:+14155238886), BOT_URL_PUBLICA (= https://robot-panaderia.onrender.com/webhook).
- **Secret File** `clave_google.json` = contenido de la llave JSON de la cuenta de servicio (pegado con Set-Clipboard).
- No van a Render: SOFFICE_PATH, ORIGEN_EXCEL, EXCEL_LOCAL_PATH (son solo de la PC).
- Construcción: 1 min 05 s la primera vez; ~30 s cuando solo cambia una variable (pasos CACHED).

### 3) Twilio
- **La cuenta de Twilio NUNCA se había creado** (era un pendiente). Creada el 1-oct-2026 con el correo de la UTN
  (el mismo de GitHub y Render). Respuestas de bienvenida: Messaging, With code.
- Sandbox: el usuario se unió desde su celular con el código del Sandbox; número **+1 415 523 8886**; webhook
  "When a message comes in" = https://robot-panaderia.onrender.com/webhook (POST). El chat muestra el nombre genérico
  "PhishGuard" (número compartido del Sandbox; no se puede cambiar).
- **Plan:** primero Limited trial (sin tarjeta, 30 días, ~100 mensajes); **la prueba gratis bloquea los mensajes enviados
  por la API** (confirmación con WA-… y recibo): error HTTP 400 "ContentSid Required" (solo deja plantillas de Twilio).
  Las respuestas inmediatas (TwiML: menú, "⏳ Recibimos tu pedido…") sí pasan. **Decisión del usuario: pasar a Full
  access** (cuenta **Individual**, perfil de cumplimiento aprobado, recarga **$20** prepago, **Auto-recharge apagado**,
  sin mensualidad). Saldo visto: $19.82. Motivo clave aceptado por el usuario: el número de orden debe llegar SÍ O SÍ al
  cliente (SINPE, retiro, comando del repartidor).
- Auth Token: cambiado una vez (token secundario → primario) tras quedar visible en una captura.

### 4) Problemas encontrados y resueltos (en orden)
1. Llave de Google borrada por accidente (el JSON se abrió en VS Code y quedó en 1 byte, guardado). **Recuperada** de la
   copia original en `FUNIVER MAESTRIA\maestria big data\proyecto final maestria\DESARROLLO pf\APIS PF robot\`
   (2387 bytes, mismo nombre); probar_conexion TODO BIEN. Lección: no abrir el JSON en VS Code; copiarlo con
   `Get-Content -Raw | Set-Clipboard`. La copia de la maestría es el respaldo de la llave.
2. Variables escritas como ".env" (el nombre del archivo, no el valor) y GOOGLE_SERVICE_ACCOUNT_FILE con la ruta de la
   PC: corregidos (/etc/secrets/clave_google.json).
3. **"Firma de Twilio invalida"**: el token de Render tenía **31 de 32 caracteres** (se copió seleccionando el texto en
   la caja angosta de Twilio). Se arregló copiando con el iconito de copiar. De paso quedó el cambio fe9b754
   (BOT_URL_PUBLICA + registro de direcciones probadas), útil si algún día cambia el proxy.
   Lección: **las claves se copian SIEMPRE con el botón de copiar**, nunca seleccionando.
4. Botón **Rollback** tocado por error: volvió a 43e2460 y Render **apagó el Auto-Deploy**. Se desplegó el último commit
   (Manual Deploy → Deploy latest commit) y se volvió a poner Auto-Deploy en On Commit (Settings → Deploy).
5. **"ContentSid Required"** (prueba gratis) → Full access (punto 3).
6. **LibreOffice se caía en Render** ("Fatal exception: Signal 6" en SvpSalInstance; Debian 12 con LibreOffice 7.4):
   ningún pedido se registraba (el bot avisaba "No pudimos registrar tu pedido"). Arreglo a4e5754: caja Ubuntu 24.04 +
   LibreOffice 24.2 (el mismo entorno donde el recálculo del v107 se probó: 3.4 s, 0 errores). Los intentos de las
   10:10 p. m. y 11:46 p. m. NO se registraron.
7. **Sueño del plan gratis**: a los 15 min sin tráfico Render apaga el servicio; el primer mensaje lo despierta
   (~1 min) y su respuesta se pierde (Twilio espera ~15 s). Pendiente: ping (#42 del mapa V17).

### 5) Mediciones (dato para la tesis)
| Qué | PC del usuario | Render gratis (0.1 CPU) |
|---|---|---|
| Escritura "registrar pedido" (Sí → confirmación WA-119) | 7.9 s | < 1 min (12:03 → 12:04 a. m.) |
| Escritura "marcar entregado" (Entregado → recibo) | 7.9 s | < 1 min (12:06 → 12:07 a. m.) |
| Respuesta inmediata (menú, "⏳") | — | al instante (1-3 s) |
| Arranque en frío (servicio dormido) | — | ~1 min; el primer mensaje se pierde |
| Construcción/despliegue | — | 1 min 05 s (primera) / ~30 s (cambio de variable) |
- Segundos exactos: se pueden sacar de los Logs de Render (líneas "Trabajo en la fila" y "guardado (intento N)").

### 6) Pedido de prueba en Drive (evidencia)
- **WA-119**: cliente = nombre de perfil de WhatsApp del usuario; 2 Combo 1 + 1 CupCake; retiro; SINPE Móvil;
  total **₡7,098.68**; ESTADO CONFIRMADO; ENTREGADO 2-oct 12:06 a. m. Escrito en Drive por el bot desde Render.
- Revisar en la próxima sesión cómo lo toma 'Flujo caja diario' (fecha 2-oct-2026) y la versión de Drive actual.

### 7) Respaldos y archivos (nada se borró)
- Escritorio\download: `paquete_hosting.zip` (**328016**), `drive_excel_ANTES_37.py` (**10797**), `app_ANTES_firma.py`
  (3508), `config_ANTES_firma.py` (2687), `Dockerfile_ANTES_lo` (1817), y los archivos descargados (Dockerfile,
  dockerignore.txt, drive_excel.py, app.py, config.py).
- **Copia local dañada a propósito por error:** a las 7:02 p. m. se ejecutó `romper_prueba.py` (botón ▷ de VS Code) y se
  respondió "si": NODO INTEGRAL C19/C20 rotas en `datos_prueba\maestro_prueba.xlsx`. Respaldo automático:
  `datos_prueba\maestro_prueba_RESPALDO_20261001_190238.xlsx`. **Restaurar antes de cualquier prueba --local.**

### 8) Ideas y notas para la tesis (pedidas por el usuario)
- **Camino a producción:** número propio de la panadería como WhatsApp Sender (verificación del negocio con Meta y nombre
  visible aprobado, en lugar de "PhishGuard"; sin "join"); fotos en el menú (imagen con texto, tarjetas/carrusel o lista
  con botones vía plantillas de contenido, o el catálogo propio de WhatsApp); el Excel como base de datos es cuello de
  botella con ~250 ventas/día (plan con más CPU o base de datos + resumen al Excel); costos por mensaje (Twilio + Meta);
  aviso de datos personales (ya en el texto de confirmación); atención humana (#41).
- Para el video/demostración: guardar el número del Sandbox como contacto ("Panadería Cañas (bot)") para que WhatsApp
  muestre ese nombre.
- **Idea futura del usuario:** reutilizar la estructura (app.py, conversation.py, fila de escrituras, Dockerfile, Render,
  Twilio) como plantilla para otros negocios con catálogo y fotos; cambiar solo el conector de datos, textos y reglas.
  Revisar la normativa de la universidad sobre reutilizar el código de la tesis.

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V15 (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V15.md`).
2. **Restaurar la copia local** desde `maestro_prueba_RESPALDO_20261001_190238.xlsx` (foto previa, Copy-Item, tamaño).
3. **Ping anti-sueño** (#42): cron-job.org (o UptimeRobot) visitando https://robot-panaderia.onrender.com/salud cada
   10 min. Ojo: 750 h gratis/mes por workspace (un servicio 24/7 ≈ 720-744 h).
4. **Seguridad (#43):** regenerar la clave de Google Maps (restringida a Routes API) y volver a cambiar el Auth Token de
   Twilio (quedó visible otra vez); actualizar .env, Render (con el botón copiar) y el libro de mapas.
5. GitHub Actions a las 7:00 CR (NODO + competencia + precios, #14); cache de competencia en Drive (#20); respaldo diario
   del Excel; endpoint de anulación solo para administración.
6. Evaluar Render Starter (#49) con los tiempos medidos; comprobar antes de la entrega que los tutores en México pueden
   unirse al Sandbox (cuenta Full access).
7. SaaS de demostración, Power BI, entrega final, trabajo escrito (sin cambios en el plan). Menores: #36, #4, #17, #19,
   #22, #21.

### Tips aprendidos hoy
- La vista del editor de VS Code no es el disco: para saber el tamaño real de un archivo, `Get-Item` en la Terminal.
- Cerrar una pestaña sin cambios no pregunta nada (Ctrl+F4); si hay un punto blanco (●) son cambios sin guardar.
- En Render: "Save, rebuild, and deploy" redespliega; "Rollback" vuelve atrás y apaga el Auto-Deploy.
- Gunicorn no anota cada visita: que no salga nada en Logs no prueba que no llegó un mensaje.
- Twilio Sandbox desconecta tras unos días sin mensajes: se vuelve a mandar el mismo "join …".

---

# CONTENIDO V14 (sin cambios)

# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V14

*Última actualización: 29-sep-2026 (noche, ~8:15 p. m. CR). Reemplaza a la V13. La V13 completa sigue abajo SIN CAMBIOS
(sección "CONTENIDO V13", que a su vez trae la V12, la V11, la V10, la V9 y la V8); lo nuevo está en este bloque
"ACTUALIZACIÓN V14". Subir este archivo y el libro de mapas V16 al inicio del próximo chat.*
**Este archivo va a GitHub: NUNCA poner claves aquí.**

*Nota de fechas: el chat empezó el 29-sep-2026 a las ~12:40 a. m. CR (pausa de madrugada) y siguió de ~12:00 m. a
~8:15 p. m. Las compras de prueba llevan fecha fija 2026-09-28. Horas de Drive en UTC (v107: 29-sep 22:26 UTC =
4:26 p. m. CR).*

---

## ACTUALIZACIÓN V14 (29-sep-2026, ~8 h activas) — Fase del bot: decisiones (Sesión A), compras de prueba, m013, marcar_entregado y bot completo con simulador

### Estado general
- 5 de 5 grupos completos (sin cambios). **Bot de WhatsApp: COMPLETO EN CÓDIGO y probado en el simulador (sin Twilio).**
  Falta conectarlo a WhatsApp real = fase de hosting (Render). Solo queda #26 del modelo (redacción del usuario).

### Commits de este chat (todos con push)
- **84041bd** — bitácora V13 (hecho antes de empezar; confirmado con `git log`).
- **848faf8** — `compra_demo_semana.py` (**5139**, raíz): compra de prueba opción B (evidencia de la tesis).
- **99118df** — `migraciones/m013_bot_telefono_entrega.py` (**13106**).
- **65e1f17** — Paso 2: `microservicios/grupo1_captura_operativa/pedidos_digitales.py` (**35856**), `catalogo.py`
  (**8648**), `probar_pedidos_digitales.py` (**12955**). 3 files, 210 insertions, 8 deletions (líneas modificadas).
- **c444327** — Paso 3: carpeta `whatsapp_bot/` (11 archivos; conversation.py **21416**). 974 insertions, 510 deletions
  (el código del bot viejo, que queda en el historial y en `download\whatsapp_bot_ANTES_p3`).
- **Último commit de código: c444327.** Pendiente al cierre: commit de esta bitácora V14 (solo este archivo).
- **Drive: versión 107** (m013). probar_conexion TODO BIEN: 25 hojas, 8 precios sin cambio (Baguette 842.69, CupCake
  343.16, Cf01 2,179.275), LibreOffice encontrado.
- **Copia local `datos_prueba\maestro_prueba.xlsx` = 249857 bytes (29-sep 4:18:28 p. m.) = compra B + m013 = igual a
  Drive v107** (mismo criterio de siempre: mismas operaciones aplicadas en los dos lados y valores verificados; no es
  copia byte a byte, por eso el archivo bajado de Drive v97 medía 247424 y la local 247427).

### 1) Decisiones del usuario (Sesión A del bot, 29-sep-2026)
| Tema | Decisión |
|---|---|
| Idioma | **Solo español.** El #16 (nombres en inglés) queda SIN EFECTO; no hay elección de idioma ni doble plantilla. |
| #33 Teléfono | **Opción B: se guarda.** Columna TELEFONO al final de TablaInterfaceDigital; CLIENTE = nombre de perfil de WhatsApp. |
| #28 Confirmación | Texto aprobado: resumen con total real (incl. envío), forma de pago, "pedido en firme / los alimentos no se reciben de vuelta" y "guardamos tu número solo para avisarte sobre tu pedido". |
| Recibo al entregar | **Opción C:** función compartida `marcar_entregado(id_pedido)`; hoy la llama el comando "ENTREGADO WA-xxx" del repartidor por WhatsApp; después, el botón del SaaS. Columna ENTREGADO nueva; **ESTADO no se toca** (Flujo y Control suman solo "CONFIRMADO"). |
| Pedido entregado | **No se anula** (regla nueva, Regla #35). |
| Canal | **Sandbox de Twilio.** Sin trámite de plantillas de Meta (el bot solo responde dentro de la ventana de 24 h que abre el cliente). El número propio queda en la tesis como camino a producción (~US$1.15/mes el número + US$0.005 por mensaje; ~US$5/mes con 50 pedidos de prueba; ojo: según varias fuentes, desde el 1-oct-2026 Meta cobra también los mensajes de servicio dentro de la ventana). |
| Datos nuevos | **Opción C:** SINPE de la panadería en NODO INTEGRAL fila 22 (número SUPUESTO 8888-8888, Regla #27); números autorizados para ENTREGADO en el `.env` (no en el Excel ni en la copia de demostración de los tutores). |
| Forma de elegir | **Menús numerados** leídos del Excel en el momento (sin plantillas de Twilio: el catálogo cambia cada día). |
| Vencimiento | Pedido sin confirmar se borra solo tras **30 minutos sin mensajes** del cliente (el reloj se reinicia con cada mensaje; solo antes del "Sí"). Regla automática, sin intervención humana (Regla #36). |
| Atención humana | **Sin opción "Hablar con alguien"** (el bot no lleva a ninguna persona). El saludo muestra directo el menú de productos. Confirmación y recibo terminan con "¿Algún problema con tu pedido? Comunícate directamente con la panadería." Conectar atención humana = mejora para producción. |
| Textos extra aprobados | Retiro: "Puedes pasar a retirarlo a la panadería; al entregártelo te enviaremos tu recibo por aquí." SINPE: "Recuerda hacer el SINPE Móvil de ₡X al 8888-8888 e indicar WA-xxx en el detalle." |
- Aclaración para la tesis: el "traductor del JSON" (#33) NO era de idioma; ahora se llama **conector del bot con el
  Excel** (`whatsapp_bot/conector_excel.py`).
- Cobro del envío (ya existía desde el 24-sep): el cliente manda el pin → `cotizar_envio` → Google Maps (distancia por
  calle) → km × Tarifa Base KM (NODO C3 = 500), sin IVA, redondeado HACIA ARRIBA al múltiplo de 50; fuera del radio =
  aviso; retiro = envío 0. Google se consulta una sola vez por pedido.

### 2) Compras de prueba (opción B, Regla #34 + #27) — Drive
- Script `compra_demo_semana.py` (usa la función real `registrar_ingresos`; fecha fija 2026-09-28; `--local` o Drive con
  "SI"). **F005 Solera** (filas 38-44): Ha01 5, Me01 3, Le01 6, Bo01 5, Ma03 1, Az01 1, Bu01 2 = ₡155,835. **F006 MaxiPalí**
  (fila 45): L01 8 = ₡8,400. Total **₡164,235**, costos congelados de la TCM (Regla #33).
- Existencias después: harina 5.77, mejorante 3.1, levadura 6.6999, bolsas 5, margarina Tumba 1.0934, azúcar 1.1063,
  leche 8.55, burbuja 2.28 (sal y manteca 0.7241 sin cambio). Alcanza para ~5 hornadas de 100 Baguette y 5 de 100 CupCake.
- Flujo caja diario 28-sep: E 155,835, F 8,400, H 164,235, J acumulado 969,522.97 → **805,287.97**. Solo 'Flujo caja
  diario' (729 fórmulas) y 'Control de Inventario' (36) leen la tabla de ingresos; **el modelo financiero (VAN, TIR, PE)
  no cambia**.
- Doble revisión: local (2a corrida DUPLICADO, archivo sin guardar de nuevo: 248719, 12:06:10 p. m.) y Drive (2a corrida
  DUPLICADO).

### 3) m013_bot_telefono_entrega — APLICADA (local y Drive v107)
- A) TablaInterfaceDigital **A7:O25 → A7:Q25**: P TELEFONO (texto), Q ENTREGADO (fecha y hora, "d-mmm hh:mm"), vacías en
  los 18 pedidos existentes. B) NODO INTEGRAL **A22:E22**: Bot WhatsApp | Número SINPE Móvil | 8888-8888 | Manual (Dato del
  negocio) | explicación. C) DiccionarioDatos filas **306-308** (OPS.WA.TELEFONO, OPS.WA.ENTREGADO, NODO.SINPE_MOVIL).
- Revisa todo antes de escribir; "Ya estaba aplicada" en la 2a corrida (local y Drive). Evento EVT-000020 (local).
- comparar_libros (local): 5,059 fórmulas iguales, 60 celdas (2 + 5 + 42 + 11 del evento), **0 valores calculados
  cambiados, 0 errores**. Probada antes también en el entorno de Claude con LibreOffice (mismo resultado).

### 4) Paso 2 — pedidos_digitales.py, catalogo.py, probar_pedidos_digitales.py (commit 65e1f17)
- `registrar_pedido_digital(..., telefono=None)`: guarda "whatsapp:+506…" limpio como **+506…** (normalizar_telefono;
  TelefonoInvalido si no tiene forma de número). Solo escribe TELEFONO si llega (sirve con copias sin m013).
- **`marcar_entregado(id_pedido, momento=None, origen="WhatsApp")`**: rechaza inexistente (LookupError), fila -A, pedido
  anulado o no CONFIRMADO (EntregaNoPermitida); si ya estaba entregado → estado YA_ENTREGADO sin escribir (no hay segundo
  recibo); si no, escribe fecha y hora en ENTREGADO + evento **ENTREGA_PEDIDO** y devuelve los datos del recibo (cliente,
  teléfono, líneas, subtotal, descuento, envío, total, forma de pago, hora). No envía WhatsApp.
- `anular_pedido_digital`: rechaza un pedido ya entregado (Regla #35).
- `catalogo.leer_reglas_del_dia`: agrega `sinpe_movil` (por etiqueta "Número SINPE Móvil"; None si falta la m013).
- Prueba en la PC con Google real: TODO OK (Cañas 1.05 km → envío ₡550; TELEFONO +50688887777; SINPE 8888-8888; WA-120
  ENTREGADO 17:15:29 con ESTADO CONFIRMADO; YA_ENTREGADO; 4 rechazos 6d-6g; flujo cuadra, diferencia 1685.38). Foto
  `maestro_prueba_ANTES_prueba_p2.xlsx` → prueba → copia restaurada (249857, 4:18:28 p. m.).

### 5) Paso 3 — Bot completo (carpeta whatsapp_bot/, commit c444327)
- Se corre como paquete desde ROBOT PANADERO: `py -m whatsapp_bot.simulador` y `py -m whatsapp_bot.app`.
- `conversation.py` (21416): la conversación (saludo + promoción NODO C19 + menú de hoy → cantidad 1-99 → agregar /
  quitar / continuar → envío con pin o retiro → Efectivo / SINPE → resumen #28 → Sí / Cambiar / Cancelar), "cancelar" en
  cualquier momento, vencimiento de 30 min (solo avisa si había productos; registra PEDIDO_ABANDONADO en el log = dato
  para una "tasa de abandono"), comando ENTREGADO tolerante a signos ("Entregado 120:", "ENTREGADO: WA-120."; rechaza
  "no entregado 120"). No habla con Twilio ni con el Excel.
- `conector_excel.py`: catálogo del día (combos activos o descuento de feriado/día de pago, productos con nombre exacto de
  la TCM, SINPE, promoción), cotizaciones, registrar, entregar; traduce errores del robot a mensajes para el cliente
  (TotalCambio → se vuelve a cotizar y se muestra el total nuevo).
- `fila_escrituras.py`: las escrituras al Excel van UNA A LA VEZ en segundo plano (Twilio espera ~15 s el webhook).
- `nucleo.py`: une conversación + conector + fila (igual para app y simulador). `app.py`: Flask; revisa la firma de
  Twilio (con arreglo para el proxy https de Render); responde al instante con TwiML; lo lento va a la fila y el
  resultado se manda por la API de Twilio; `/salud` para Render. En Render: `gunicorn --workers 1 --threads 4
  whatsapp_bot.app:app` (un solo proceso: las conversaciones viven en memoria).
- `simulador.py`: chat en la Terminal contra la copia local (comandos /ubicacion canas|liberia|lat,lng, /repartidor,
  /cliente, /esperar N, /salir). `config.py`: lee el `.env` de la raíz (sin claves ni teléfonos reales).
  `setup_content_templates.py`: OBSOLETO (se conserva como historia). `.env.example` y `README.md` actualizados.
- Prueba en la PC: pedido WA-120 (2 Combo 4 + 1 CupCake, retiro, SINPE, total ₡9,723.68) y entrega con recibo.
  **Tiempo medido: 7.9 s por escritura (registro y entrega) en la PC.** Estimado en Render gratis: 30 s a 1.5 min
  (medir). El cliente recibe primero un aviso inmediato (1-3 s) y luego la confirmación con el número de pedido.
- Hallazgo corregido: "entregado 120:" no se reconocía → regex tolerante.

### 6) Hallazgos anotados (mapa V16)
- **#36 (nuevo, pendiente):** margarina Ma01: costo 1,000 por caja de 4 y precio G12 = costo/4 × 1.05 = 262.50 por barra,
  pero 'Control de Inventario' resta compras (10) y ventas (3) como la misma unidad (mismo tipo de hueco que tuvo el
  café; Regla #32 = inventario en unidad de venta).
- **#37 (nuevo, pendiente para hosting):** Drive sube versiones por su cuenta segundos después de cada subida:
  compras 98 → 99/100 (2 reintentos en F006, entró en el 3er y último intento de editar_excel) → 101 → 103 sin escribir;
  m013 una subida (104) → 107. Con escrituras seguidas (bot + SaaS) podrían fallar registros legítimos. Ver también #31.
- Precios del catálogo: el menú muestra el precio con 2 decimales (Cf01 2,179.275 → ₡2,179.28); el total que se cobra
  siempre sale de cotizar_pedido (igual al registrado).

### 7) Libro de mapas V16 (LIBRO_MAPA_DE_CELDAS_PFV16.xlsx) — FUERA DEL REPO (tiene claves, Regla #15)
- Reglas Globales: **Regla #35** (pedido entregado: ENTREGADO sin tocar ESTADO; no se anula) y **Regla #36** (el bot decide
  solo sobre lo que no tocó el Excel; las personas sobre lo registrado; 30 min automáticos). Decisiones Pendientes:
  #16, #28, #33 RESUELTOS (|| V16); #31 con evidencia nueva; bloque "ACTUALIZACIÓN V16" con #36, #37 (pendientes) y
  #38 recibo (C), #39 canal Sandbox, #40 compras de prueba, #41 atención humana (producción). Mapas 'Interface digital' y
  'NODO INTEGRAL' con bloque "ACTUALIZACIÓN V16". Pestaña "Cambios V16".

### Respaldos y fotos (nada se borró)
- Escritorio\download: `compra_demo_semana.py` (5139), `m013_bot_telefono_entrega.py` (13106), `pedidos_digitales_ANTES_p2.py`
  (28132), `catalogo_ANTES_p2.py` (8140), `probar_pedidos_digitales_ANTES_p2.py` (9092), carpeta `whatsapp_bot_ANTES_p3`
  (7 archivos del bot viejo, sin .env), `whatsapp_bot.zip` (el que se armó para el chat), `whatsapp_bot_p3.zip` (21221),
  `conversation_ANTES_regex.py` (21317).
- datos_prueba: `maestro_prueba_ANTES_compraB.xlsx` (247427 = Drive v97), `maestro_prueba_ANTES_m013.xlsx` (248719 = compra
  B), `maestro_prueba_ANTES_prueba_p2.xlsx` (249857 = compra B + m013 = copia actual).

### Pendientes (próximo chat y después)
1. Commit de esta bitácora V14 si no quedó hecho (`git add RESPALDO_PROYECTO_ROBOT_PANADERIA_V14.md`).
2. **Fase de hosting (próximo chat):**
   - Activar el Sandbox de WhatsApp en Twilio (unirse con el código); en el `.env` y en Render: TWILIO_ACCOUNT_SID,
     TWILIO_AUTH_TOKEN, TWILIO_WHATSAPP_FROM, **BOT_NUMEROS_AUTORIZADOS** (el usuario los escribe; nunca en el chat) y las
     de Google Drive (en Render: GOOGLE_SERVICE_ACCOUNT_JSON).
   - **Render necesita LibreOffice** para recalcular (editar_excel): probablemente con Docker (imagen con LibreOffice).
     Decidir y probar. Servicio web con gunicorn 1 worker; URL del webhook `https://<servicio>/webhook` en el Sandbox;
     probar con el celular; **medir tiempos reales en Render** (dato para la tesis).
   - GitHub Actions a las 7:00 CR (NODO + competencia + precios, #14); cache de competencia en Drive (#20); respaldo
     diario del Excel; endpoint de anulación solo para administración; revisar #37 (esperar unos segundos entre
     escrituras seguidas o más intentos).
3. SaaS de demostración (copia del Excel, usuario y contraseña, sin tocar el maestro) + botón "Entregado" que llama
   `marcar_entregado`. Entrega final (ZIP con `git archive` desde etiqueta v1.0-tesis, `.env.example` sin claves, link
   para los tutores, video de respaldo). Trabajo escrito (normativa + #26).
4. Menores: #36 (margarina), #4, #17, #19, #22, #21 (clave de Gemini en el mapa, acción del usuario).

### Tips aprendidos hoy
- Un "git commit" sin "git add" antes responde "no changes added to commit" y no crea nada.
- Git cuenta como "deletions" las líneas MODIFICADAS (la versión vieja), no solo las borradas.
- `Expand-Archive -Force` sobre ROBOT PANADERO reemplaza los archivos de la carpeta que trae el ZIP.
- Si un comando de copia por tamaño no hace nada, revisar primero si el archivo nuevo sí está en download.
- comparar_libros con 0 valores cambiados + 0 errores = la migración solo agregó lo anunciado.

---

# CONTENIDO V13 (sin cambios)

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
