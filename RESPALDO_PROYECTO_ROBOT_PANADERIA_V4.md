# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V4

*Última actualización: 24-sep-2026 (cierre de sesión). Reemplaza a la V3. Guarda este archivo en la carpeta del proyecto
(junto a CLAUDE.md) y súbelo al inicio de cualquier chat nuevo con Claude.*

**Nota para quien retome este proyecto (incluido Claude en un chat nuevo):**
el usuario (Francisco) es principiante total en Git, terminal y programación.
Necesita instrucciones muy concretas y paso a paso, con nombres exactos de
botones y menús. Los comandos de Terminal (escribir texto y dar Enter) le
funcionan mejor que buscar iconitos con el mouse. Claude actúa como **tutor**:
explicar qué se construye y por qué, no solo dar pasos.

---

## 1. Qué es el proyecto

Estudio de factibilidad de una panadería (Cañas, Guanacaste, Costa Rica) con
un Excel maestro (financiero, 24 hojas) + robot en Python que lo mantiene
actualizado desde un SaaS y un bot de WhatsApp. Proyecto final de la Maestría
en Marketing Digital y Big Data (UNINI-MX, tutores en México). Todo debe
funcionar en la nube y de forma automática, sin depender de la PC del usuario.

## 2. Arquitectura confirmada (no cambiar sin razón)

1. **Conexión a Excel**: Google Drive + Cuenta de Servicio, sin login humano.
   - Correo: `robot-panaderia@robot-panaderia-canas.iam.gserviceaccount.com`
   - Proyecto Google Cloud: `robot-panaderia-canas` (Google Drive API habilitada)
   - `EXCEL_MAESTRO_FILE_ID=1uOY84vrtjmV_YvReHWxcHlumEATpCnlw`
   - Mecanismo: descargar `.xlsx` → editar con `openpyxl` → recalcular con
     LibreOffice → subir. Se mantiene como `.xlsx` real.
   - En Compartir > engranaje debe estar marcado "Editors" en "People who can
     download, copy, and print" (sin eso: error 403 `cannotDownloadFile`).
   - Se descartó Microsoft Graph/OneDrive (login ocasional; el M365 Developer
     Program ya no da tenant gratis a cuentas nuevas).
2. **Bot de WhatsApp**: Twilio. Pedido por selección guiada (Quick Reply /
   List Picker), nunca texto libre.
3. **Hosting planeado**: Render (bot, siempre activo, escrituras en tiempo
   real) + GitHub Actions (solo tareas programadas).
4. **Una sola fuente de verdad**: NODO INTEGRAL es el único proceso que llama
   APIs externas (clima, feriados, tipo de cambio). SaaS y bot solo leen.
5. **No existe "cliente único"**: todo en unidades y pedidos.
6. **`openpyxl.insert_rows()` no ajusta fórmulas**: el robot escribe
   explícitamente el texto de cada fórmula en las filas nuevas.
7. **Power BI**: después de que haya datos reales (Grupos 1-2), vía Google
   Drive para Escritorio (sincronización local). NUNCA con enlace público.
8. **El Excel maestro NUNCA se abre con Google Sheets** (Sheets quitó el
   prefijo `_xlfn.` a TEXTJOIN en NODO INTEGRAL C19/C20 → `#NAME?`). Solo se
   toca con el robot o descargándolo/subiéndolo a propósito. Las versiones
   "Unavailable account" en Drive son del robot (válidas).
9. **Tipo de cambio** (NODO INTEGRAL C7) = referencia oficial BCCR leída de la
   API de Hacienda (coincide al céntimo con el BCCR).
10. **Escrituras en tiempo real** (~250 ventas/día): bot y SaaS deben escribir
    por la MISMA función compartida (`escribir_pedido_en_excel()`) para no
    pisarse.
11. **Los mapas de celdas son la base del código**: todo cambio al Excel se
    refleja en el libro de mapas siguiendo el protocolo del proyecto.

## 3. Reglas de negocio

- **Captura en SaaS**: no se digitan nombres; se elige con autocompletado.
  El robot valida el código recibido como red de seguridad.
- **Canales**: despacho (mostrador) y venta en línea (bot) son ingresos
  separados, se cobran y rebajan de inventario por separado.
- **Combos**: iguales en ambos canales; solo se venden combos ACTIVOS (los
  activa su detonador, ej. temperatura) y ya traen el descuento en el precio.
  En pedidos de WhatsApp un combo es UNA línea (ej. "Combo 2").
- **Días de pago**: 15 y 30 de cada mes (en febrero, el último día); si cae
  sábado/domingo cuenta ese día. → 5% en todo, sin combos.
- **Feriados**: solo los 9 de pago obligatorio de Costa Rica (no 2 ago,
  31 ago, 1 dic). → 5% en todo, sin combos.
- **Hornadas**: una hornada registrada por el panadero suma la producción al
  inventario y rebaja la materia prima (hoja "Consumo de Materia Prima").
  Anular = poner 0 en G2 (celda naranja junto a "ORDEN ANULADA") en
  'costo pan baguette' / 'costo cupcakes'; se deja fila de corrección negativa.
- **Precio**: celda "Precio de referencia" (la digita la administración) en
  las hojas de costo. Todo cálculo de INGRESOS (proyecciones, punto de
  equilibrio, TIR, flujo caja proyecto B9/C9) usa el precio de referencia;
  los COSTOS son dinámicos (costo real por hornada). Producir más = más
  ganancia neta, no precio más bajo.
- Conversiones en TABLA CRUCE MAESTRA MP: Mejorante 3 kg, Café 250 g x 3;
  CupCake lleva 25 g de harina. Costeo de la leche en cupcakes: pendiente
  (APIs/scraping).
- En Historico Diario se mantienen los nombres "pan baguette" y "cupcakes".

## 4. Orden de construcción (5 grupos)

| Grupo | Hojas | Estado |
|---|---|---|
| 1 — Captura operativa | Ventas despacho, Interface digital, Detalle comanda digital, Ingreso de Inventario | ✅ Completo y probado |
| 2 — Inventario | Control de Inventario, Historico Diario (+ Consumo de Materia Prima) | ✅ Completo, migraciones aplicadas al Excel real |
| 3 — Integraciones | NODO INTEGRAL, TABLA CRUCE MAESTRA, TABLA CRUCE MAESTRA MP, Tabla Combos Dinamicos | 🔶 3A completo; 3B parte 1 completa; 3B parte 2 y 3C pendientes |
| 4 — Costeo | costo pan baguette, costo cupcakes, planilla, gasto operacional, valor activos | ⏳ Sin código |
| 5 — Financiero | Flujo caja diario/anual/proyecto, Punto de Equilibrio, TIR VAR, proyeccion ventas, proyeccion anual ventas | ⏳ Sin código |

## 5. Carpeta en la PC

Ruta real: `C:\Users\franf\Desktop\ROBOT PANADERO` (con espacio → usar
comillas en comandos). Terminal: PowerShell dentro de VS Code, con la
extensión "Claude Code for VS Code".

```
ROBOT PANADERO/
├── CLAUDE.md
├── .env                  (ignorado por Git; incluye GOOGLE_MAPS_API_KEY)
├── .gitignore            (UTF-8)
├── credenciales/         (clave JSON de la cuenta de servicio, ignorada)
├── romper_prueba.py      (herramienta de prueba del escudo _xlfn)
├── datos_prueba/         (maestro_prueba.xlsx, copia local fija, ignorada)
├── conexion_drive/       (drive_excel.py, recalculo.py, probar_conexion)
├── migraciones/          (m001_grupo2, m002_consumo_mp)
├── microservicios/
│   ├── grupo1_captura_operativa/   ✅
│   ├── grupo2_inventario/          ✅
│   ├── grupo3_integraciones/       🔶 (nodo_integral, google_maps)
│   ├── grupo4_costeo/              (vacía)
│   └── grupo5_financiero/          (vacía)
└── whatsapp_bot/         (conversación probada; aún no escribe a Excel)
```

**Archivos descargados**: nunca en Descargas. Van a una carpeta intermedia del
Escritorio (ej. `Escritorio\download` o `Escritorio\Prueba grupo 3`) y de ahí
se mueven con comandos de Terminal a su carpeta final.

## 6. Entorno Python

- Python 3.14. `pip` no está en PATH → usar siempre `py -m pip ...` y
  `py -m <modulo>`.
- Librerías de `requirements.txt` instaladas. LibreOffice instalado.
- Comandos útiles:
  - Probar conexión: `py -m conexion_drive.probar_conexion`
  - NODO INTEGRAL: `py -m microservicios.grupo3_integraciones.nodo_integral`
  - Google Maps (solo lee): `py -m microservicios.grupo3_integraciones.google_maps <lat> <lng>`
  - Modo prueba local (solo esa ventana de Terminal):
    `$env:ORIGEN_EXCEL="local"; $env:EXCEL_LOCAL_PATH="datos_prueba/maestro_prueba.xlsx"`
    y para volver a Drive: `Remove-Item Env:ORIGEN_EXCEL, Env:EXCEL_LOCAL_PATH`
  - Las migraciones se prueban primero con `--local` (sobre datos_prueba).

## 7. Git / GitHub

- Repo privado: https://github.com/Gigo1965/robot-panaderia.git (rama main,
  remote "origin", login vía Git Credential Manager).
- `.gitignore` protege: `.env`, `credenciales/`,
  `robot-panaderia-canas-*.json`, `*.xlsx`, `datos_prueba/`, `__pycache__`,
  `venv`. Verificado: el `.env` nunca se commiteó.
- Commits clave: "V4" (conexion_drive + grupo1), cd7db42 (Grupo 2),
  c52e289 (NODO INTEGRAL v3, tipo de cambio Hacienda), 67d3361 (escudo _xlfn
  + bitácora V3), 5becbcc (probar_conexion: 25 hojas), a3a69f1 (Grupo 3B parte 1).
- Git en 3 pasos: `git add .` (meter en la caja) → `git commit -m "..."` (cerrar
  y etiquetar) → `git push` (enviar a GitHub). Revisar antes con `git status`.
- **Regla**: nunca editar en la web de GitHub; todo en VS Code →
  `git add` / `git commit` / `git push`. README para tutores: pendiente.

## 8. Documentos fuera del repo

- **Excel maestro**: en Drive (versión local de referencia: v68).
- **Libro de mapas de celdas**: V8 (Reglas #12 no Google Sheets, #13 tipo de
  cambio; pestaña "Cambios V8"). El Excel maestro tiene 25 hojas.
- **Regla del usuario**: ninguna clave o dato de API se borra de los mapas sin
  consultarle antes. La clave de Google Maps NO vuelve a los mapas: el usuario
  la guarda en su PC y en su Drive; el robot la lee del `.env`.

## 9. Dónde quedamos (24-sep-2026, cierre)

**Hecho hoy:**
- Escudo `_xlfn` COMPLETO: repara también fórmulas MATRICIALES (NODO
  INTEGRAL C19/C20 son ArrayFormula). Probado en la PC rompiendo la copia de
  prueba a propósito (`romper_prueba.py`, que guarda
  `datos_prueba/maestro_prueba_RESPALDO.xlsx`) → el robot la reparó solo.
- `probar_conexion.py` ahora espera 25 hojas (probado contra Drive).
- Grupo 3B parte 1: `microservicios/grupo3_integraciones/google_maps.py`
  (Routes API computeRouteMatrix). Prueba real OK: 2.32 km, 7 min, envío ₡1,200.

**Decisiones Grupo 3B (proyecto académico):**
- Origen = 10.42895, -85.09425 (aprox. 100 m del local). Radio máximo 10 km
  (fuera de zona → el bot avisa). Envío = km × Tarifa Base KM (NODO C3),
  sin IVA, redondeado HACIA ARRIBA al múltiplo de ₡50. Origen, radio y
  redondeo configurables en `.env` (ORIGEN_LAT, ORIGEN_LNG, RADIO_MAXIMO_KM,
  REDONDEO_ENVIO).
- Km del pedido en su fila de Interface digital (N = DISTANCIA_KM,
  I = ENVIO_MAPS); NODO C9/C10 = solo el último pedido (referencia).
- La función compartida de pedidos es la ÚNICA puerta a Google Maps
  (excepción a la regla de NODO INTEGRAL: la distancia depende del cliente).
- Bot BILINGÜE (español/inglés), a construir en la fase de integración real
  del bot: idioma elegido al saludar con botones; el Excel no cambia.
  Pendiente decidir cómo se traducen los nombres de productos.

**Siguientes pasos:**
1. Confirmar la prueba "fuera de zona" (Liberia):
   `py -m microservicios.grupo3_integraciones.google_maps 10.6346 -85.4377`
2. Grupo 3B parte 2: conectar `google_maps.py` a `registrar_pedido_digital`
   (escribir N e I en Interface digital y NODO C9/C10). Subir al chat el
   archivo de pedidos digitales de grupo1 para integrarlo.
3. Actualizar libro de mapas a V9 con los cambios del 3B (sin la clave).
4. Grupo 3C: scraping Palí / Solera + promo Musmanni con visión de Claude.
5. Grupos 4 (costeo) y 5 (financiero).
6. Integración real del bot de WhatsApp (bilingüe, función compartida).

**Pendientes de hosting (sin decidir)**:
- Plan gratis de Actions en repo privado = 2,000 min/mes; un ping a Render
  cada 10-14 min lo excede. Opciones: cron-job.org / UptimeRobot, correr
  NODO INTEGRAL dentro de Render, repo público, o pagar minutos.
- Respaldo diario del Excel: opción preferida = marcar la versión del cierre
  como "Keep forever" vía Drive API. Ojo: Drive guarda máx. 100 versiones
  temporales (con ~250 escrituras/día cubre menos de un día) y las fijadas
  tendrían un tope (~200 por archivo, por verificar) → liberar las antiguas.

## 10. Para continuar en un chat nuevo

Subir: este archivo, el Excel maestro más reciente (descargado de Drive, sin
abrirlo en Sheets) y el libro de mapas más reciente. Indicar desde qué punto
de la sección 9 se retoma.
