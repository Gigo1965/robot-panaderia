# Reporte de la migración m003 — Grupo 3C (proveedores MaxiPalí / Solera)

*24-sep-2026 (noche). Script: `migraciones/m003_proveedores_3c.py`. Excel de origen: el descargado de Drive el 24-sep 17:45 (224,159 bytes). Opción confirmada por el usuario: **D1** (nombres + conversiones + precios verificados).*

## 1. Resumen

| Bloque | Qué | Celdas |
|---|---|---|
| A | Leche Suli → Leche Semidescremada Dos Pinos -1 L | 11 |
| B | 7 materias primas de Solera reemplazadas (MP, Ingreso, Control) | 21 |
| C | Conversiones col E de TABLA CRUCE MAESTRA MP | 6 |
| D | Precios verificados 24-sep (MP D10:D20, TCM D10, D11, D12, D14) | 15 |
| E | Notas `\|\| m003` en DiccionarioDatos (se agregó texto, no se borró nada) | 5 |
| — | Evento en RegistroEventos (fila 9, EVT-000005) | 1 fila |
| **Total** | | **58 celdas + 1 evento** |

## 2. Doble revisión

**Revisión 1 (dentro del script):** antes de escribir revisó las 58 celdas contra el valor "antes" esperado. Si una sola no coincide, se detiene sin cambiar nada; se probó con un archivo alterado a propósito y no se modificó ni un byte. Después de escribir volvió a leer cada celda. Una segunda corrida responde "Ya estaba aplicada". Ningún texto nuevo empieza con "=" (solo las 6 fórmulas de la col E).

**Revisión 2 (script independiente que compara TODO el libro celda por celda):**

- Celdas distintas entre el original y el nuevo: 69 = 58 anunciadas + 11 del evento en RegistroEventos.
- Cambios NO anunciados: **ninguno**. Anunciados que faltan: **ninguno**. Con valor distinto al anunciado: **ninguno**.
- 25 hojas, mismo orden. Tablas: solo `tblRegistroEventos` crece de A4:L8 a A4:L9 (el evento). Validaciones de datos, formato condicional y celdas combinadas: **iguales**.
- NODO INTEGRAL C19/C20 conservan `_xlfn.TEXTJOIN`.
- Recálculo con LibreOffice: **4,538 fórmulas, 0 errores** (igual que el original: 4,538 y 0).

## 3. Cambios celda por celda (antes → después)

| # | Bloque | Hoja | Celda | Antes | Después |
|---|---|---|---|---|---|
| 1 | A | TABLA CRUCE MAESTRA | C10 | Leche Suli Semidescremada 1.5% grasa  - 1 L | Leche Semidescremada Dos Pinos -1 L |
| 2 | A | Control de Inventario | A7 | Leche Suli Semidescremada 1.5% grasa  - 1 L | Leche Semidescremada Dos Pinos -1 L |
| 3 | A | Ingreso de Inventario | C4 | Leche Suli Semidescremada 1.5% grasa  - 1 L | Leche Semidescremada Dos Pinos -1 L |
| 4 | A | Ingreso de Inventario | C7 | Leche Suli Semidescremada 1.5% grasa  - 1 L | Leche Semidescremada Dos Pinos -1 L |
| 5 | A | Ventas despacho | E22 | Leche Suli Semidescremada 1.5% grasa  1 L | Leche Semidescremada Dos Pinos -1 L |
| 6 | A | Ventas despacho | E42 | Leche Suli Semidescremada 1.5% grasa  1 L | Leche Semidescremada Dos Pinos -1 L |
| 7 | A | Ventas despacho | E44 | Leche Suli Semidescremada 1.5% grasa  1 L | Leche Semidescremada Dos Pinos -1 L |
| 8 | A | Ventas despacho | E147 | Leche Suli Semidescremada 1.5% grasa  1 L | Leche Semidescremada Dos Pinos -1 L |
| 9 | A | Tabla Combos Dinamicos | D5 | 3 Baguette + 1 Leche Suli | 3 Baguette + 1 Leche Semidescremada Dos Pinos -1 L |
| 10 | A | Tabla Combos Dinamicos | K5 | 🔥 Combo 1 activo: 3 Baguette + 1 Leche Suli | 🔥 Combo 1 activo: 3 Baguette + 1 Leche Semidescremada Dos Pinos -1 L |
| 11 | A | DiccionarioDatos | N80 | Descripcion del contenido del combo, redactada para que el robot/bot de WhatsApp la use al armar el pedido (ej. '3 Baguette + 1 Leche Suli'). | Descripcion del contenido del combo, redactada para que el robot/bot de WhatsApp la use al armar el pedido (ej. '3 Baguette + 1 Leche Semidescremada Dos Pinos -1 L'). |
| 12 | B | TABLA CRUCE MAESTRA MP | C10 | MEJORANTE STAR PAN PLUS 3 KILOS | MEJORANTE PANNIVITA PRO-PAN 1 KG HOMO-PLUS |
| 13 | B | Ingreso de Inventario | C10 | MEJORANTE STAR PAN PLUS 5 KILOS | MEJORANTE PANNIVITA PRO-PAN 1 KG HOMO-PLUS |
| 14 | B | Control de Inventario | A12 | MEJORANTE STAR PAN PLUS 5 KILOS | MEJORANTE PANNIVITA PRO-PAN 1 KG HOMO-PLUS |
| 15 | B | TABLA CRUCE MAESTRA MP | C13 | SAL FINA DIAMANTE 25 KG COLOR VERDE. GRANO NORMAL | SAL SARU COMUN ECONOMICA 25 KG SACO BRIN |
| 16 | B | Ingreso de Inventario | C13 | SAL FINA DIAMANTE 25 KG COLOR VERDE. GRANO NORMAL | SAL SARU COMUN ECONOMICA 25 KG SACO BRIN |
| 17 | B | Control de Inventario | A15 | SAL FINA DIAMANTE 25 KG COLOR VERDE. GRANO NORMAL | SAL SARU COMUN ECONOMICA 25 KG SACO BRIN |
| 18 | B | TABLA CRUCE MAESTRA MP | C15 | LEVADURA TEMPO SALADA 500 GR | LEVADURA FRESCA GOLONDRINA 400 GR CIP |
| 19 | B | Ingreso de Inventario | C15 | LEVADURA TEMPO SALADA 500 GR | LEVADURA FRESCA GOLONDRINA 400 GR CIP |
| 20 | B | Control de Inventario | A17 | LEVADURA TEMPO SALADA 500 GR | LEVADURA FRESCA GOLONDRINA 400 GR CIP |
| 21 | B | TABLA CRUCE MAESTRA MP | C16 | MERMELADA GUAYABA F3 1 KG ANGEL MANGA | MERMELADA GUAYABA HORNEO 5 KG CCR CAJA C2 (FRUTA TRITURADA) |
| 22 | B | Ingreso de Inventario | C16 | MERMELADA GUAYABA F3 1 KG ANGEL MANGA | MERMELADA GUAYABA HORNEO 5 KG CCR CAJA C2 (FRUTA TRITURADA) |
| 23 | B | Control de Inventario | A18 | MERMELADA GUAYABA F3 1 KG ANGEL MANGA | MERMELADA GUAYABA HORNEO 5 KG CCR CAJA C2 (FRUTA TRITURADA) |
| 24 | B | TABLA CRUCE MAESTRA MP | C17 | BOLSAS PAPEL KRAFT 1 BAGUETTE 500 UNDS BIO | BOLSAS PAPEL KRAFT 2 BAGUETTE 100 UNDS BIODEGRADABLE |
| 25 | B | Ingreso de Inventario | C17 | BOLSAS PAPEL KRAFT 1 BAGUETTE 500 UNDS BIO | BOLSAS PAPEL KRAFT 2 BAGUETTE 100 UNDS BIODEGRADABLE |
| 26 | B | Control de Inventario | A19 | BOLSAS PAPEL KRAFT 1 BAGUETTE 500 UNDS BIO | BOLSAS PAPEL KRAFT 2 BAGUETTE 100 UNDS BIODEGRADABLE |
| 27 | B | TABLA CRUCE MAESTRA MP | C18 | BURBUJA #6 BIO 50 UND. 6″X6″ POSTRE, REPOSTERIA BAZAGO  | BURBUJA FOAM #6 125 UNID POSTRE BL |
| 28 | B | Ingreso de Inventario | C18 | BURBUJA #6 BIO 50 UND. 6″X6″ POSTRE, REPOSTERIA BAZAGO  | BURBUJA FOAM #6 125 UNID POSTRE BL |
| 29 | B | Control de Inventario | A20 | BURBUJA #6 BIO 50 UND. 6″X6″ POSTRE, REPOSTERIA BAZAGO  | BURBUJA FOAM #6 125 UNID POSTRE BL |
| 30 | B | TABLA CRUCE MAESTRA MP | C19 | CAFE NARANJO PURO 250 GR X 3 UNIDADES | Café Naranjo puro 250gr x 10 Unidades |
| 31 | B | Ingreso de Inventario | C19 | CAFE NARANJO PURO 250 GR X 10 UNIDADES | Café Naranjo puro 250gr x 10 Unidades |
| 32 | B | Control de Inventario | A21 | CAFE NARANJO PURO 250 GR X 10 UNIDADES | Café Naranjo puro 250gr x 10 Unidades |
| 33 | C | TABLA CRUCE MAESTRA MP | E10 | =D10/3000 | =D10/1000 |
| 34 | C | TABLA CRUCE MAESTRA MP | E15 | =D15/500 | =D15/400 |
| 35 | C | TABLA CRUCE MAESTRA MP | E16 | =D16/1000 | =D16/5000 |
| 36 | C | TABLA CRUCE MAESTRA MP | E17 | =D17/500 | =D17/100 |
| 37 | C | TABLA CRUCE MAESTRA MP | E18 | =D18/50 | =D18/125 |
| 38 | C | TABLA CRUCE MAESTRA MP | E19 | =D19/750 | =D19/2500 |
| 39 | D | TABLA CRUCE MAESTRA MP | D10 | 5000 | 1815 |
| 40 | D | TABLA CRUCE MAESTRA MP | D11 | 25000 | 17510 |
| 41 | D | TABLA CRUCE MAESTRA MP | D12 | 24000 | 16490 |
| 42 | D | TABLA CRUCE MAESTRA MP | D13 | 25000 | 3310 |
| 43 | D | TABLA CRUCE MAESTRA MP | D14 | 25000 | 28845 |
| 44 | D | TABLA CRUCE MAESTRA MP | D15 | 1000 | 735 |
| 45 | D | TABLA CRUCE MAESTRA MP | D16 | 1000 | 4140 |
| 46 | D | TABLA CRUCE MAESTRA MP | D17 | 1000 | 2400 |
| 47 | D | TABLA CRUCE MAESTRA MP | D18 | 1000 | 7275 |
| 48 | D | TABLA CRUCE MAESTRA MP | D19 | 2500 | 20755 |
| 49 | D | TABLA CRUCE MAESTRA MP | D20 | 3600 | 15390 |
| 50 | D | TABLA CRUCE MAESTRA | D10 | 800 | 1050 |
| 51 | D | TABLA CRUCE MAESTRA | D11 | 100 | 2300 |
| 52 | D | TABLA CRUCE MAESTRA | D12 | 750 | 1000 |
| 53 | D | TABLA CRUCE MAESTRA | D14 | 100 | 840 |
| 54 | E | DiccionarioDatos | N70 | (texto existente, sin cambios: 1118 caracteres) | mismo texto + agregado: «\|\| m003 (24-sep-2026): LAS URLs ANTERIORES YA NO SON VALIDAS (Solera re-sincronizo su catalogo en sep-2026; 8 de 11 ya no existen) y 7 productos se reemplazaron por precio y funcionalidad con aprobacion del usuario. URLs VIGENTES (SKU entre parentesis): Me01 (0549) -> https://solera.co.cr/producto/mejorante-pannivita-pro-pan-1-kg-homo-plus/ ; Ha01 (0016) -> https://solera.co.cr/producto/harina-flores-25-kg-fuerte/ ; Az01 (0021) -> https://solera.co.cr/producto/azucar-dona-maria-2kg-24-kg/ ; Sa01 (0995) -> https://solera.co.cr/producto/sal-saru-comun-economica-25-kg-saco-brin/ ; Ma02 (0030) -> https://solera.co.cr/producto/manteca-max-pan-bulk-25-kgs/ ; Le01 (0137) -> https://solera.co.cr/producto/levadura-fresca-golondrina-400-gr-cip/ ; Me02 (0957) -> https://solera.co.cr/producto/mermelada-guayaba-horneo-5-kg-ccr-caja-c2-fruta-triturada/ ; Bo01 (0516) -> https://solera.co.cr/producto/bolsas-papel-kraft-2-baguette-100-unds-biodegradable/ ; Bu01 (0179) -> https://solera.co.cr/producto/burbuja-foam-6-125-unid-postre-bl/ ; Ca01 (0474) -> https://solera.co.cr/producto/cafe-naranjo-puro-250gr-x-10-unidades/ ; Ma03 (0034) -> https://solera.co.cr/producto/margarina-tumba-numar-7-5-kilos-60-barritas-x125gr/ . Nunca usar 'SAL ROCA 25 KG' (no es para consumo humano).» |
| 55 | E | DiccionarioDatos | N72 | (texto existente, sin cambios: 465 caracteres) | mismo texto + agregado: «\|\| m003 (24-sep-2026): METODO ACTUALIZADO por decision del usuario. Principal = API publica de catalogo de cada tienda (gratis, sin IA, no depende del diseno): Solera (WooCommerce Store API) https://solera.co.cr/wp-json/wc/store/v1/products?slug=<slug> , precio en prices.price (currency_minor_unit 0); MaxiPali (VTEX) https://www.maxipali.co.cr/api/catalog_system/pub/products/search?fq=alternateIds_Ean:<EAN> , precio en items[0].sellers[0].commertialOffer.Price. Respaldo = ScrapeGraphAI con Google Gemini (clave GEMINI_API_KEY en el .env). Precios 1 vez al dia; si falla, NO se toca la celda.» |
| 56 | E | DiccionarioDatos | L72 | (texto existente, sin cambios: 13 caracteres) | mismo texto + agregado: «\|\| m003: API de catalogo (principal) + ScrapeGraphAI/Gemini (respaldo)» |
| 57 | E | DiccionarioDatos | N76 | (texto existente, sin cambios: 900 caracteres) | mismo texto + agregado: «\|\| m003 (24-sep-2026): en Ingreso de Inventario (C10, C19) y Control de Inventario (A12, A21) seguian los nombres viejos '5 KILOS' y 'X 10 UNIDADES', distintos de la tabla MP (esas compras salian con precio 0 y como 'Compras Externas Netas'). La m003 puso el mismo nombre nuevo en las 3 hojas. Ver COST.MP.FUENTE para las URLs vigentes.» |
| 58 | E | DiccionarioDatos | N232 | (texto existente, sin cambios: 337 caracteres) | mismo texto + agregado: «\|\| m003 (24-sep-2026): D10 = 1050, precio verificado en MaxiPali el 24-sep-2026 (EAN 7441001601064, Leche Semidescremada Dos Pinos 1 L). Respaldo si no hay: Sabemas UHT 1 L (EAN 7441078217694). Ya no es placeholder; el robot lo actualizara a diario.» |

Evento agregado: RegistroEventos fila 9 → `EVT-000005 · MIGRACION · m003_proveedores_3c · Proveedores 3C (MaxiPali/Solera) · OK`.

## 4. Efectos en los valores calculados

Las fórmulas no se tocaron; estos valores cambian porque cambian sus datos.

| Dónde | Antes | Después | Por qué |
|---|---|---|---|
| Control de Inventario D7 / F7 (leche) | 0 vendidas / existencia 30 | 11 vendidas / existencia 19 | Ventas despacho E22/E42/E44/E147 ya calzan con el nombre |
| Ingreso de Inventario E10 / L10 (mejorante) | 0 / Compras Externas Netas | 1,815 / Materia Prima | El nombre ya coincide con MP |
| Ingreso de Inventario E19 / L19 (café) | 0 / Compras Externas Netas | 20,755 / Materia Prima | Igual |
| TCM D8 — costo Baguette | 483.30 | 411.23 | Harina, azúcar y sal más baratas |
| TCM D9 — costo CupCake | 217.19 | 229.45 | Margarina, burbuja y leche más caras |
| TCM G8 / G9 — precio de referencia | 842.69 / 343.16 | **sin cambio** | Regla #11 (precio de referencia) |
| TCM G10 — precio leche | 840.00 | 1,102.50 | D10 real 1,050 × 1.05 |
| TCM G11 / G12 / G14 — queso, margarina, natilla | 105 / 196.88 / 105 | 2,415 / 262.50 / 882 | Queso y natilla tenían costo de relleno ₡100 |
| Combo 1 / 3 / 4 (E5 / E7 / E8) | 3,115.26 / 1,621.84 / 2,380.26 | 3,377.76 / 2,398.84 / 4,690.26 | Suben leche, natilla y queso |
| Flujo caja diario L12 | −782,246.93 | −1,754,366.93 | Las compras de prueba (mayo) ahora usan precios reales |
| Punto de Equilibrio B29 | 65,055 | 57,810 | Baja el costo del Baguette |
| NODO INTEGRAL C19 (mensaje) | …1 Leche Suli… | …1 Leche Semidescremada Dos Pinos -1 L… | Lee Combos K5 |

Lo que se actualiza solo, sin escribirlo: Consumo de Materia Prima col C y F, composición de combos C17 y la lista auxiliar de Ventas despacho L4.

## 5. Para revertir

El original está intacto en `Escritorio\respaldo_antes_m003`. Para aplicar solo nombres y conversiones (opción D2) sobre una copia de ese original: `py -m migraciones.m003_proveedores_3c --archivo "<copia>.xlsx" --sin-precios`.
