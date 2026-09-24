# Respaldo del Proyecto — Robot Panadería (Estudio de Factibilidad) — V3

*Última actualización: 24-sep-2026. Reemplaza a las versiones anteriores
(21-sep y V2 del 23-sep). Guarda este archivo en la carpeta del proyecto
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
| 3 — Integraciones | NODO INTEGRAL, TABLA CRUCE MAESTRA, TABLA CRUCE MAESTRA MP, Tabla Combos Dinamicos | 🔶 3A completo; 3B y 3C pendientes |
| 4 — Costeo | costo pan baguette, costo cupcakes, planilla, gasto operacional, valor activos | ⏳ Sin código |
| 5 — Financiero | Flujo caja diario/anual/proyecto, Punto de Equilibrio, TIR VAR, proyeccion ventas, proyeccion anual ventas | ⏳ Sin código |

## 5. Carpeta en la PC

Ruta real: `C:\Users\franf\Desktop\ROBOT PANADERO` (con espacio → usar
comillas en comandos). Terminal: PowerShell dentro de VS Code, con la
extensión "Claude Code for VS Code".

```
ROBOT PANADERO/
├── CLAUDE.md
├── .env                  (ignorado por Git)
├── .gitignore            (UTF-8)
├── credenciales/         (clave JSON de la cuenta de servicio, ignorada)
├── datos_prueba/         (maestro_prueba.xlsx, copia local fija, ignorada)
├── conexion_drive/       (drive_excel.py, recalculo.py, probar_conexion)
├── migraciones/          (m001_grupo2, m002_consumo_mp)
├── microservicios/
│   ├── grupo1_captura_operativa/   ✅
│   ├── grupo2_inventario/          ✅
│   ├── grupo3_integraciones/       🔶 (nodo_integral)
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
  - Las migraciones se prueban primero con `--local` (sobre datos_prueba).

## 7. Git / GitHub

- Repo privado: https://github.com/Gigo1965/robot-panaderia.git (rama main,
  remote "origin", login vía Git Credential Manager).
- `.gitignore` protege: `.env`, `credenciales/`,
  `robot-panaderia-canas-*.json`, `*.xlsx`, `datos_prueba/`, `__pycache__`,
  `venv`. Verificado: el `.env` nunca se commiteó.
- Commits clave: "V4" (conexion_drive + grupo1), cd7db42 (Grupo 2),
  c52e289 (NODO INTEGRAL v3, tipo de cambio Hacienda).
- **Regla**: nunca editar en la web de GitHub; todo en VS Code →
  `git add` / `git commit` / `git push`. README para tutores: pendiente.

## 8. Documentos fuera del repo

- **Excel maestro**: en Drive (versión local de referencia: v68).
- **Libro de mapas de celdas**: V8 (Reglas #12 no Google Sheets, #13 tipo de
  cambio; clave de Google Maps retirada del libro y guardada en `.env`;
  pestaña "Cambios V8").

## 9. Dónde quedamos (24-sep-2026)

**En curso — escudo `_xlfn`**: función `reparar_prefijos_xlfn()` en
`conexion_drive/recalculo.py` (repara TEXTJOIN, CONCAT, IFS, MAXIFS, MINIFS,
SWITCH sin prefijo), llamada automáticamente dentro de `_descargar()` en
`drive_excel.py` (cubre `leer_excel()` y `editar_excel()`). Falta la prueba
real: correr `romper_prueba.py` (raíz del proyecto) para dañar una fórmula a
propósito y luego `py -m conexion_drive.probar_conexion` para confirmar que
se repara sola.

**Siguientes pasos**:
1. Confirmar escudo `_xlfn` y subirlo a GitHub.
2. Grupo 3B: Google Maps (distancia / costo de envío por pedido).
3. Grupo 3C: scraping Palí / Solera + promo Musmanni con visión de Claude.
4. Grupos 4 (costeo) y 5 (financiero).
5. Integración real del bot de WhatsApp (función compartida con el SaaS).

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
