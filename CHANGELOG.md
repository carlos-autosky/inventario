# Changelog — Sistema de Inventario


## v3.38 (2026-04-07 11:15 GMT-5)

### Tema claro — arranque correcto
- `dark_mode` inicializa en `False` — la app siempre arranca en tema claro Autosky.
- Toggle 🌙 sigue funcionando para cambiar a oscuro en runtime.

### CSS tema oscuro — implementado correctamente
- `_DARK_CSS` separado que sobreescribe variables CSS y selectores cuando el toggle está activo.
- `_ACTIVE_CSS = _CSS + (_DARK_CSS if dark else "")` — inyección dinámica por rerun.
- El CSS oscuro cubre: fondo app, sidebar, inputs, selects, dropdowns, file uploader, expanders, métricas, tablas HTML internas, download buttons.

### Sidebar versión dinámica
- Logo del sidebar usaba `v3.36` hardcodeada — corregido a `{APP_VERSION}` y `{BUILD_TIME}`.
- Nunca más desfasado con la versión real del archivo.

### main_web.py — basura eliminada
- `config.toml` suelto al final del archivo (residuo de edición anterior) eliminado.



## v3.35 (2026-04-07 00:00 GMT-5)

### Análisis de Período — tarjetas eliminadas
- Eliminadas "Ventas período" y "Margen período" de las KPI cards. Solo quedan: Item más rotación, Item más vendido, Item más rentable.

### Plantilla Toma Física — RESUMEN GENERAL automático
- **Flujo anterior**: llenar hojas individuales + llenar resumen manualmente (doble trabajo).
- **Flujo nuevo**: el usuario solo llena la columna CANTIDAD en cada hoja de ubicación. El RESUMEN GENERAL se calcula **automáticamente** con fórmulas Excel.
- Fórmula por celda de ubicación: `=IFERROR('NombreHoja'!D{fila}, 0)` — sin macros, sin VBA, funciona en cualquier versión de Excel y LibreOffice.
- Columna TOTAL del resumen: `=SUM(D{ri}:X{ri})` automático.
- Fila TOTAL GENERAL: suma automática de todas las filas.
- Celdas del resumen en verde muy pálido (#F0FDF4) para distinguirlas visualmente de las editables.
- Nota de advertencia en fila 2: "NO editar este resumen".
- Si se agrega una nueva hoja en el Excel: la nueva ubicación se crea automáticamente al importar.

## v3.32 (2026-04-06 23:24 GMT-5)

### Toma Física — Importación Excel: lógica de hoja eliminada corregida
- Comportamiento anterior: ubicación ausente del Excel → sin cambios.
- Comportamiento nuevo: ubicación presente en el sistema pero **ausente del Excel** → todos sus ítems pasan a `qty = 0`.
  Razón: si el usuario eliminó esa hoja del Excel, significa que esa ubicación fue inspeccionada y no tiene unidades de ningún SKU.
- El mensaje de confirmación ahora informa cuántos ítems fueron puestos en cero.
- Ubicaciones nuevas en el Excel → se crean automáticamente (sin cambios).

### Toma Física — Columna anterior en pantalla (confirmación)
- Ya existía desde versiones anteriores: columna "Anterior (fecha)" en gris no editable (`tk.Label`) que muestra la cantidad de la última toma guardada para cada SKU en cada ubicación. Solo aparece si hay historial previo.

## v3.31 (2026-04-06 23:21 GMT-5)

### Toma Física — Excel colores pasteles
- Paleta completamente clara: cabeceras amarillo pastel, datos gris muy suave (#F1F5F9), filas zebra blanco/#F8FAFC, totales verde pastel, anterior violeta pastel. Solo el título mantiene azul marino (#1E3A5F).
- El archivo se abre automáticamente al exportar.

### Toma Física — Importación desde Excel (nuevo)
- Botón **📥 Importar** lee el Excel exportado previamente.
- **Hoja nueva** en el Excel → crea la ubicación automáticamente.
- **Hoja eliminada** del Excel → la ubicación permanece intacta en el sistema (no se borra, no se pone en 0).
- Solo se actualizan las ubicaciones presentes en el archivo importado.

### Kardex — nombre de archivo incluye SKU
- Si hay filtro activo: `kardex_ACPAA-2.pdf`, `kardex_ACPAA-2.xlsx`, `kardex_ACPAA-2.html`.
- Sin filtro: `kardex.pdf`, `kardex.xlsx`, `kardex.html`.

### Kardex — auto-abrir archivos exportados
- PDF, Excel y HTML se abren automáticamente en el programa asociado al guardar.

### Kardex — autocomplete SKU ejecuta búsqueda automática
- Al seleccionar un SKU del listbox predictivo, si hay fechas cargadas, ejecuta el Kardex inmediatamente sin necesidad de presionar "Generar".

## v3.30 (2026-04-06 23:11 GMT-5)

### Rotación y Compras — paneles expandibles corregidos
- `panels.grid_rowconfigure(0, weight=1)` corregido (era row=1 — los paneles están en row=0).
- Los escenarios Marítimo y Aéreo ahora ocupan todo el espacio vertical disponible.

### Kardex — autocomplete SKU — selección con mouse/Enter/Tab
- Clic, Enter y Tab ahora seleccionan correctamente la sugerencia.
- Teclas ↑ y ↓ navegan por la lista desde el campo de entrada.
- Escape cierra el listbox sin seleccionar.
- `_pick_mouse` usa `after(50ms)` para que el clic se registre antes de leer la selección.

### Kardex — EGRESO color gris (sin rojo)
- Pantalla: tag `kdx_egr` mantiene celeste/azul (ya sin rojo).
- HTML: EGRESO y EGR DEV.PROV → fondo `#F3F4F6` / texto `#374151` (gris oscuro).
- PDF: `C_EGR` → `#F3F4F6`.
- Excel: fill EGR → `E5E7EB`.
- Leyenda HTML actualizada con borde para distinguir del fondo blanco.

## v3.29 (2026-04-06 22:56 GMT-5)

### Kardex — mejoras visuales y funcionales
- **Colores pastel azul/celeste**: paleta monocromática azul, sin rojo/verde/naranja intensos.
  INICIO=azul oscuro | INGRESO=azul claro | EGRESO=celeste | TRA=azul medio | DEV=azul suave.
- **Fila SUBTOTAL por SKU**: al cambiar de SKU se inserta fila acumulada en azul pastel fuerte.
- **Texto predictivo SKU**: al escribir en el campo SKU aparece listbox flotante con sugerencias de códigos y nombres. Selección con clic o Enter. Se cierra al perder el foco.
- **Exportación HTML**: botón 🌐 HTML genera reporte con fondo blanco, colores pastel por tipo, fila SUBTOTAL, leyenda de colores, se abre automáticamente en el navegador.
- **Auto-inserción de "/" en fechas**: `_make_date_entry()` formatea automáticamente al tipear solo números (ddmmyyyy → dd/mm/yyyy).

## v3.28 (2026-04-06 22:44 GMT-5)

### Nueva pestaña — Kardex de Inventario
- Filtros: rango de fechas (DESDE/HASTA) + SKU (código o nombre, vacío = todos).
- Botones: Generar (pantalla) | PDF | Excel.
- **Algoritmo `_calc_kardex()`**:
  1. Calcula stock y costo promedio ponderado acumulados hasta `d_from - 1 día` (saldo inicial).
  2. Inserta fila **SALDO INICIAL** por SKU con el stock y costo vigente al inicio del período.
  3. Recorre cada movimiento del período en orden cronológico:
     - **INGRESO** (ING+FAC): suma al saldo, actualiza costo promedio ponderado.
     - **ING DEV.CLI** (ING+NCT): suma al saldo, costo promedio NO cambia.
     - **EGRESO** (EGR+FAC): resta del saldo, muestra costo promedio vigente.
     - **EGR DEV.PROV** (EGR+NCT): resta del saldo, costo NO cambia.
     - **TRANSFERENCIA** (TRA): saldo consolidado = 0 (sale de una bodega, entra a otra). Costo NO cambia.
  4. Columnas: Fecha | Código | Nombre | Referencia | Descripción | Tipo Mov. | Cantidad | V.Unit | Costo Prom. | Saldo Uds | Valor Inv.
- **Colores en pantalla**: azul=INICIO, verde=INGRESO, rojo=EGRESO, violeta=TRA, amarillo=DEV.
- **PDF**: landscape A4, colores de fondo pastel por tipo, texto wrap en Nombre/Descripción.
- **Excel**: formato profesional, colores por tipo, número_format contable, paneles congelados en fila 4.

### Análisis de Período — barra DESDE/HASTA alineada a la izquierda.

## v3.26 (2026-04-06 22:31 GMT-5)

### Costo Promedio Ponderado — algoritmo corregido
- Eliminado costo 0: fallback a `Valor Unitario` del movimiento, luego `PVP × 0.6`.
- Promedio ponderado real: `(qty_ant × cp_ant + qty_nueva × c_nueva) / (qty_ant + qty_nueva)`.
- Costo calculado por fila de venta (no por SKU total), usando el costo vigente hasta esa fecha.
- Gráfica mensual ahora refleja el costo real de lo vendido cada mes.

### Rotación y Compras — dos escenarios separados
- **Panel izquierdo (🚢 Marítimo)**: lead time configurable (default 45d), borde teal.
- **Panel derecho (✈ Aéreo)**: lead time configurable (default 15d), borde azul.
- Fórmulas corregidas:
  - `Consumo/día = ventas_u / días_período`
  - `P.Reorden(u) = consumo_día × lead_time`
  - `Sug.Compra = max(0, consumo_día × (lead_time + stock_seg) − stock_disp)`
  - `Días Inv. = stock_disp / consumo_día`
  - Estado CRÍTICO si `días_inv < lead_time` (no hardcoded 15d).
- Barra de contexto con explicación completa de cada fórmula.

### HTML Muestras — KPI color
- Total Enviadas y Total Devueltas en gris (#6B7280). Saldo en Cliente mantiene rojo/verde.

## v3.25 (2026-04-06 22:12 GMT-5)

### Reporte Muestras PDF — hoja horizontal, sin solapamiento
- Cambiado a `landscape(A4)` (29.7×21cm horizontal).
- Anchos de columna calculados dinámicamente sobre el ancho útil de la página.
- Texto largo en "Nombre Producto" y "Descripción" usa `Paragraph` con wrap automático — elimina solapamiento.
- KPI cards se distribuyen a lo ancho de la página.

### Reporte Muestras HTML — números en gris legible
- Celdas numéricas (`td.num`) usan `color:#6B7280` con `font-variant-numeric:tabular-nums` para mejor legibilidad sin perder contraste.

### Histórico de Compras — subtotales y costo promedio
- Filas ordenadas por Código Producto → Fecha.
- Al final de cada grupo SKU se inserta una fila **SUBTOTAL** (azul/amarillo) con:
  - Cantidad total acumulada del SKU
  - Valor Total acumulado
  - **Costo Promedio Ponderado**: calculado progresivamente compra a compra `(costo_nuevo + costo_anterior) / 2`
- Fila **TOTAL GENERAL** al final con totales globales.
- Con filtro activo: misma lógica, solo sobre los SKUs filtrados.

## v3.24 (2026-04-06 21:59 GMT-5)

### Reporte Muestras — PDF y HTML fondo blanco
- Rediseño completo: fondo blanco, texto negro, encabezados azul oscuro (#1E3A5F), alternancia de filas gris muy claro.
- Saldo en rojo (#DC2626) si hay unidades pendientes, verde (#059669) si saldado.
- Campo **Descripción** agregado en la tabla de detalle de movimientos (PDF y HTML).

### Análisis de Periodo — Costo Promedio Ponderado
- Nuevo algoritmo: recorre compras históricas ordenadas por fecha por SKU.
  - Primera compra: costo_prom = valor_unitario de esa compra.
  - Compras siguientes: costo_prom = (costo_actual + costo_anterior) / 2.
- El costo del periodo se calcula como: unidades_vendidas × costo_promedio_SKU.
- Tabla mensual muestra columna "Ventas(Egreso)" y "Costo" basado en costo promedio.
- Top 10 rentabilidad usa costo promedio por SKU (columna "Costo Prom.").

### Nueva pestaña — Histórico de Compras
- Muestra todas las facturas de compra (ING+FAC) del archivo cargado.
- Columnas: Fecha, Factura, Código, Nombre Producto, **Descripción**, Cantidad, Valor Total, **Valor Unitario** (= Valor Total / Cantidad).
- Filtro de texto en tiempo real por código o nombre de producto.
- Se actualiza automáticamente al ejecutar el análisis.

### Scrollbar separada del gráfico en Análisis de Periodo
- `_make_analysis_panel` envuelve el tree en su propio frame para que el scrollbar vertical quede contenido y no se superponga al canvas del gráfico.

## v3.23 (2026-04-06 21:41 GMT-5)

### Muestras por Cliente — Reporte ejecutivo PDF y HTML
- Nuevo selector de cliente en barra superior de la pestaña.
- Botón **PDF**: genera reporte ejecutivo con portada, KPIs (enviadas/devueltas/saldo), tabla de resumen por SKU con última fecha de devolución, y detalle completo de movimientos ordenado por producto y fecha. Saldo en rojo si tiene unidades pendientes, verde si está saldado.
- Botón **HTML**: genera el mismo reporte en HTML con tema oscuro, se abre automáticamente en el navegador. Listo para imprimir o compartir.
- `_get_sample_report_data()`: construye el dataset desde `r.filtered` usando `is_sample_out` / `is_sample_in` — coherente con el engine.
- El selector de clientes se actualiza automáticamente al ejecutar el análisis.

## v3.22 (2026-04-06 21:29 GMT-5)

### SKU x Bodega — exportación Excel corregida
- Causa del fallo: conversión `int(fv) if fv == int(fv)` fallaba silenciosamente con valores NaN o enteros grandes, dejando celdas vacías sin avisar.
- Reescrito `export_pivot_excel` con `try/except` externo que muestra el error real en pantalla y en el log.
- Columnas numéricas normalizadas con `pd.to_numeric(...).fillna(0)` antes de escribir, eliminando NaN.
- Mensaje de confirmación muestra cantidad de SKU y bodegas exportadas.

## v3.21 (2026-04-06 21:19 GMT-5)

### SKU x Bodega — simplificación y corrección de exportación
- Eliminado selector "Tipo de valor" — la pestaña opera exclusivamente en modo "Stock neto por bodega".
- Barra de controles simplificada: solo "Generar reporte" y "Exportar Excel".
- `export_pivot_excel` corregido: ya no referencia `pivot_mode` (que fue eliminado). Título del Excel fijo: "STOCK NETO POR BODEGA".
- `run_pivot_report` usa modo fijo "Cantidad neta por bodega" sin depender de widget eliminado.

## v3.20 (2026-04-06 20:44 GMT-5)

### SKU x Bodega — filtros corregidos
- **Por defecto excluye la Bodega Principal** — el reporte muestra solo bodegas externas (muestras, clientes, sucursales) sin necesidad de configuración.
- **Filtro BODEGAS RPT aplicado**: si el usuario selecciona bodegas específicas en el selector BODEGAS RPT, el pivot respeta esa selección exacta.
- Filas con TOTAL = 0 excluidas automáticamente del resultado.
- Log indica las bodegas incluidas cuando son 3 o menos.

## v3.19 (2026-04-06 20:34 GMT-5)

### SKU x Bodega — cálculo corregido
- `run_pivot_report()` en modo "Cantidad neta por bodega" ahora usa directamente `r.inventory_by_warehouse` (la misma fuente que las tarjetas KPI), eliminando la discrepancia de cálculo.
- Los filtros de fecha, bodega y SKU excluidos se aplican correctamente porque `inventory_by_warehouse` ya viene filtrado del engine.
- Modos de movimiento (ventas/compras/TRA) usan `r.filtered` que también tiene los filtros aplicados.

### Carga automática de Excel para pruebas
- Al iniciar, si existe `C:\\Users\\carlo\\Downloads\\inventario_movimientos_consolidado.xlsx`, se carga automáticamente sin mostrar diálogo.
- Implementado en `_auto_load()` — se puede eliminar fácilmente cuando no sea necesario.

## v3.18 (2026-04-06 20:23 GMT-5)

### Auto-fit de columnas (doble clic en separador de heading)
- Doble clic en el borde derecho de cualquier encabezado ajusta el ancho automáticamente al contenido más ancho (heading o datos), igual que Excel.
- Implementado en `_make_tree_in` mediante `tree.bind("<Double-Button-1>", _autofit_col)`.

### Pestaña SKU x Bodega — encabezados limpios
- La palabra "Bodega" eliminada de todos los títulos de columna-bodega en `_fill_pivot`.
- Ejemplo: "Bodega Principal" → "Principal", "Bodega Muestras Lima" → "Muestras Lima".

### Timestamp GMT-5 corregido
- `BUILD_TIMESTAMP` ahora usa la hora real GMT-5 al momento de generar la versión.

### Nombre de carpeta en ZIP
- A partir de esta versión, la carpeta interna del ZIP lleva el nombre de la versión completa (ej. `v03_18`).

## v3.15 (2026-04-07 02:50 GMT-5)

### two_line — ajuste línea 1999
- `"Valor Compras (ING)"` → `"Compras $ (Ingreso)|"` (sin espacios extra alrededor del pipe)

## v3.13 (2026-04-07 02:20 GMT-5)

### Encabezados de columnas — redefinición completa según especificación
- `two_line` actualizado con los labels exactos definidos por el usuario:
  `Cód.Prod` | `Nomb. Prod` | `Categ Prod` | `Compras $` | `Ventas $` | `Compras $ (Ingreso)` | `Ventas $(Egreso)` | `Inventario ($)` | `Inventario` | `Unitario $ Promedio` | `Stock $` | `Dev.Prov` | `N/C Cliente` | `Muestras Enviadas` | `Muestras Devueltas` | `Stock Disponible` | `Stock en Muestras` | `Stock Total` | `Stock en Cliente`

## v3.12 (2026-04-07 02:00 GMT-5)

### Encabezados de tabla — segunda línea visible
- Causa raíz: `padding=(4, 10)` en `Dark.Treeview.Heading` dejaba solo 10px vertical, insuficiente para mostrar 2 líneas de texto de 10px bold.
- Corregido a `padding=(6, 16)` — 16px vertical permite que ambas líneas del encabezado se rendericen completamente.
- Eliminado style muerto `Sep.Treeview.Heading` (residuo de experimentos anteriores).
- `_apply_zoom` actualiza el padding del heading proporcionalmente al font size.

## v3.11 (2026-04-07 01:40 GMT-5)

### Detalle por SKU — Movimiento de Unidades — encabezados corregidos
- Causa raíz identificada: `_col_w()` asignaba 82px a columnas como `Dev_Proveedor`, `Muestras_Enviadas`, etc., truncando la segunda línea del encabezado.
- `_col_w()` actualizado con keywords para todos los tipos de columna de movimiento:
  `Dev. al Proveedor` → 110px | `Dev. del Cliente` → 240px | `Muestras Enviadas/Devueltas` → 110px | `Stock Disponible/en Muestras` → 110px | `Compras/Ventas` → 105px
- `fill_tree` ahora usa `max(_col_w(nombre_interno), _col_w(label_visible))` para garantizar que el ancho sea suficiente independientemente de cómo se llame la columna internamente.

## v3.10 (2026-04-07 01:15 GMT-5)

### Detalle por SKU — Movimiento de Unidades (revert limpio)
- Eliminados todos los experimentos de separadores (canvas, SEP_cols, sub-filas por grupo) sin dejar código muerto.
- `_fill_sku_unit_grouped` y `_paint_sep_headings` eliminados completamente.
- `_fill_sku_split` simplificado: usa `fill_tree` estándar para ambas ventanas (Valores Financieros y Movimiento de Unidades).
- Encabezados de columnas en `fill_tree` actualizados a texto completo:
  `Compras (Ingreso)` | `Dev. al Proveedor` | `Ventas (Egreso)` | `Dev. del Cliente` | `Muestras Enviadas` | `Muestras Devueltas` | `Stock Disponible` | `Stock en Muestras` | `Stock Total`

## v3.9 (2026-04-07 00:50 GMT-5)

### Detalle por SKU — Movimiento de Unidades
- **Separadores**: canvas eliminado del área de datos (era estático). Nuevo método `_paint_sep_headings()` dibuja líneas blancas de 2px SOLO sobre el heading (área fija, no scrollea), y se repinta al soltar el mouse (`<ButtonRelease-1>`) y durante el arrastre (`<B1-Motion>`), siguiendo al usuario cuando expande columnas.
- Columnas SEP reducidas a 3px con carácter `|` en datos — siempre alineado con la posición real de la columna.
- **Encabezados texto completo** verificados: `Código Producto`, `Nombre Producto`, `Compras (Ingreso)`, `Dev. al Proveedor`, `Ventas (Egreso)`, `Dev. del Cliente`, `Muestras Enviadas`, `Muestras Devueltas`, `Stock Disponible`, `Stock en Muestras`, `Stock Total`.

## v3.13 (2026-04-07 02:20 GMT-5)

### Encabezados de columnas — redefinición completa según especificación
- `two_line` actualizado con los labels exactos definidos por el usuario:
  `Cód.Prod` | `Nomb. Prod` | `Categ Prod` | `Compras $` | `Ventas $` | `Compras $ (Ingreso)` | `Ventas $(Egreso)` | `Inventario ($)` | `Inventario` | `Unitario $ Promedio` | `Stock $` | `Dev.Prov` | `N/C Cliente` | `Muestras Enviadas` | `Muestras Devueltas` | `Stock Disponible` | `Stock en Muestras` | `Stock Total` | `Stock en Cliente`

## v3.12 (2026-04-07 02:00 GMT-5)

### Encabezados de tabla — segunda línea visible
- Causa raíz: `padding=(4, 10)` en `Dark.Treeview.Heading` dejaba solo 10px vertical, insuficiente para mostrar 2 líneas de texto de 10px bold.
- Corregido a `padding=(6, 16)` — 16px vertical permite que ambas líneas del encabezado se rendericen completamente.
- Eliminado style muerto `Sep.Treeview.Heading` (residuo de experimentos anteriores).
- `_apply_zoom` actualiza el padding del heading proporcionalmente al font size.

## v3.11 (2026-04-07 01:40 GMT-5)

### Detalle por SKU — Movimiento de Unidades — encabezados corregidos
- Causa raíz identificada: `_col_w()` asignaba 82px a columnas como `Dev_Proveedor`, `Muestras_Enviadas`, etc., truncando la segunda línea del encabezado.
- `_col_w()` actualizado con keywords para todos los tipos de columna de movimiento:
  `Dev. al Proveedor` → 110px | `Dev. del Cliente` → 240px | `Muestras Enviadas/Devueltas` → 110px | `Stock Disponible/en Muestras` → 110px | `Compras/Ventas` → 105px
- `fill_tree` ahora usa `max(_col_w(nombre_interno), _col_w(label_visible))` para garantizar que el ancho sea suficiente independientemente de cómo se llame la columna internamente.

## v3.10 (2026-04-07 01:15 GMT-5)

### Detalle por SKU — Movimiento de Unidades (revert limpio)
- Eliminados todos los experimentos de separadores (canvas, SEP_cols, sub-filas por grupo) sin dejar código muerto.
- `_fill_sku_unit_grouped` y `_paint_sep_headings` eliminados completamente.
- `_fill_sku_split` simplificado: usa `fill_tree` estándar para ambas ventanas (Valores Financieros y Movimiento de Unidades).
- Encabezados de columnas en `fill_tree` actualizados a texto completo:
  `Compras (Ingreso)` | `Dev. al Proveedor` | `Ventas (Egreso)` | `Dev. del Cliente` | `Muestras Enviadas` | `Muestras Devueltas` | `Stock Disponible` | `Stock en Muestras` | `Stock Total`

## v3.9 (2026-04-07 00:50 GMT-5)

### Detalle por SKU — Movimiento de Unidades (rediseño completo)
- Eliminados canvas overlay y columnas SEP — eran frágiles ante scroll y resize.
- Nueva estrategia: **4 sub-filas por SKU**, una por grupo, cada una con su color de fondo nativo de Treeview:
  - Azul    `#0F2744` → Compras (Ingreso) | Dev. al Proveedor
  - Verde   `#0F2D20` → Ventas (Egreso)   | Dev. del Cliente
  - Violeta `#1E1533` → Muestras Enviadas | Muestras Devueltas
  - Teal    `#0D2626` → Stock Disponible  | Stock en Muestras | Stock Total
- Zebra interna: tono ligeramente más claro en filas impares de cada SKU.
- Encabezados: texto completo en 2 líneas, sin abreviaciones.
- El separador es el propio cambio de color de fondo entre grupos — nativo, sigue scroll y resize sin canvas.

## v3.8 (2026-04-07 00:20 GMT-5)

### Detalle por SKU — Movimiento de Unidades
- Eliminadas líneas rojas en cabecera (canvas overlay removido de headings).
- Separadores verticales ahora son **líneas blancas de 2px** que atraviesan el grid completo (heading + filas de datos), dibujadas via `_draw_unit_sep_lines()` con `tk.Canvas` posicionado absolutamente sobre el tree. Se actualizan al redimensionar la ventana.
- Encabezados con **texto completo** sin abreviaciones:
  `Código Producto` | `Nombre Producto` | `Compras (Ingreso)` | `Dev. al Proveedor` | `Ventas (Egreso)` | `Dev. del Cliente` | `Muestras Enviadas` | `Muestras Devueltas` | `Stock Disponible` | `Stock en Muestras` | `Stock Total`

## v3.7 (2026-04-06 23:55 GMT-5)

### Detalle por SKU — Movimiento de Unidades (corrección v3.6)
- Encabezados ahora únicos y descriptivos en todas las columnas:
  `Código Prod.` | `Nombre Producto` | `Compras (ING)` | `Dev. Proveedor` | `Ventas (EGR)` | `Dev. Cliente` | `Muest. Enviadas` | `Muest. Devueltas` | `Stk. Disponible` | `Stk. Muestras` | `Stk. Total`
- Separadores verticales: columnas SEP_n de 12px con carácter `║` en heading y datos.
- Canvas overlay `_draw_sep_lines()` dibuja líneas rojas (#DC2626) sobre los headings de las cols separadoras (workaround a limitación de ttk.Treeview que no permite color por columna).

## v3.6 (2026-04-06 23:30 GMT-5)

### Detalle por SKU — Movimiento de Unidades
- Separadores verticales corregidos: columnas `SEP_n` de 8px con heading rojo oscuro (#7F1D1D), visibles como línea divisoria real entre grupos.
- Encabezados de columna ahora son únicos y descriptivos (sin repetición):
  - `Compras (ING)` | `Dev. Proveedor` | `Ventas (EGR)` | `Dev. Cliente` | `Muest. Enviadas` | `Muest. Devueltas` | `Stk. Disponible` | `Stk. Muestras` | `Stk. Total`

### Zoom In / Zoom Out — todos los grids
- Barra de zoom (−  +  ↺) agregada encima de cada Treeview mediante `_make_tree_in()`.
- Controla `font size` (rango 7–18px) y `rowheight` proporcional en todos los grids simultáneamente.
- `_apply_zoom()` actualiza el `ttk.Style` global y los tags de fuente bold.
- Botón ↺ resetea al tamaño por defecto (11px).

## v3.5 (2026-04-06 23:00 GMT-5)

### Detalle por SKU — Movimiento de Unidades
- Reemplazados separadores horizontales (filas azules) por **separadores verticales** entre grupos de columnas.
- Se inyectan columnas angostas (6px) con carácter │ entre cada grupo:
  - Compras | Dev. Proveedor  **│**  Ventas | Dev. Cliente  **│**  Muestras Enviadas | Muestras Devueltas  **│**  Stock Disponible | Stock Muestras | Stock Total
- La lógica de renderizado queda en `_fill_sku_unit_grouped()`.

## v3.4 (2026-04-06 22:30 GMT-5)

### Pestaña Detalle por SKU

#### Valores Financieros
- Eliminada columna "Categoría Producto".
- Renombradas las 3 columnas de valor para identificar claramente su origen:
  - `Valor_Compras` → **Valor Compras (ING)**
  - `Valor_Ventas`  → **Valor Ventas (EGR)**
  - `Valor Inventario` → **Valor Inventario ($)**

#### Movimiento de Unidades
- Eliminada columna "Categoría Producto".
- Columnas agrupadas visualmente con filas separadoras de color azul:
  - **── COMPRAS / DEV. PROVEEDOR ──** → Compras | Dev. Proveedor
  - **── VENTAS / DEV. CLIENTE ──**   → Ventas | Dev. Cliente
  - **── MUESTRAS ──**                → Muestras Enviadas | Muestras Devueltas
  - **── STOCK ──**                   → Stock Disponible | Stock Muestras | Stock Total
- Nuevo método `_fill_sku_unit_grouped()` maneja el renderizado por grupos.

## v3.3 (2026-04-06 22:00 GMT-5)

### Toolbar — reordenamiento
- Nuevo orden: Cargar Excel | Toma Fisica | Plantilla Toma | Guardar Config | Exportar Excel | BASE (label)
- Label BASE movido al extremo derecho como referencia contextual.

### Filters — reordenamiento
- Nuevo orden: CORTE | BODEGAS | BODEGAS RPT | SKU EXCL

### Corrección de cierre limpio (bgerror al salir)
- Registrado `protocol("WM_DELETE_WINDOW", self._on_closing)` en `__init__`.
- Nuevo método `_on_closing()`: cancela el `after()` pendiente del debounce antes de destruir la ventana.
- `_schedule_recalc()` ahora guarda el ID del `after()` en `self._after_id`.
- `_do_recalc()` limpia `self._after_id` al ejecutarse.
- Sin cambios en lógica de negocio ni layout.

## v3.2 (2026-04-06 16:13 GMT-5)

### Inventario por Bodega
- Encabezados de las columnas de valor aclarados para distinguir "Valor Unit. Promedio" de "Valor Stock".
- Anchos de columnas recalculados según el máximo número de caracteres visible en filas y encabezados, para evitar textos truncados.
- Columna "Nombre Producto" sin truncado manual, ampliada dinámicamente según el texto más largo dentro de un límite visual razonable.
- Mayor padding vertical en encabezados para mejorar la visualización de títulos en 2 líneas.

## v3.1 (2026-04-06 15:40 GMT-5)

### Inventario por Bodega
- Encabezados de columnas ajustados para visualizarse correctamente en 2 líneas dentro del contenedor de títulos
- Columna "Nombre Producto" ampliada y truncada con elipsis para evitar desbordes visuales manteniendo legibilidad


## v001.3 (2026-04-05)

### Correcciones UI / Tablas
- Fuente de tablas aumentada a 11px (Segoe UI) — más legible
- Anchos de columna compactos y proporcionales por tipo semántico (nombre, bodega, número, etc.)
- Headers de Treeview en una sola línea, sin cortes
- Números siempre alineados a la derecha; texto a la izquierda
- Filas zebra (alternancia de color) en todas las tablas

### Inventario por Bodega
- Filas con stock = 0 eliminadas del resultado
- Coloración visual por grupo: azul = Disponible, verde = Muestras/Otras bodegas
- Tabla de toma física: verde = coincide, rojo = diferencia

### Interacción
- Cambiar filtro de bodegas o checkboxes dispara recálculo automático (300ms debounce)
- Botón "Aplicar" junto a fecha de corte para recálculo explícito
- `run_analysis(silent=True)` para recálculo silencioso sin mensajes de error molestos

### SKU
- Exclusión de SKU aplicada correctamente antes del análisis
- Columna renombrada "Stock Muestras" (sin espacio "en") para evitar headers largos

### Muestras
- Clientes Activos filtra correctamente saldo > 0 (sin cambios en lógica de engine)
- Tabla Muestras por Cliente con zebra y columnas compactas

### Layout / Diseño
- Tarjetas KPI más bajas (menor pady, fuente 15px en valor)
- KPIs distribuidos en 2 filas de 6/5, sin espacios vacíos
- Toolbar más compacto (altura 34px, íconos en botones)
- Panel de filtros más ajustado en altura
- Scrollable frame de bodegas más pequeño (68px visible)

### Persistencia
- Config se guarda automáticamente en cada análisis
- `config_state` se actualiza tras cada `persist_config` para coherencia interna

## v001.2 (2026-04-05)
- Versión base funcional con toda la lógica de negocio

## v002.9.2 (2026-04-06 13:34:55 GMT-5)
- Corregido error de inicio en versiones debug/rebuild causadas por referencias a métodos o paneles no creados.
- Reajustado panel FILTERS a barra compacta alineada a la izquierda, sin expansión vertical innecesaria.
- Controles de Fecha Corte, Bodegas, SKU Excluidos y Bodegas Reporte agrupados en secuencia a la izquierda.

## v3.0 (2026-04-06 15:10:00 GMT-5)
- Base oficial retomada desde v002.9.2 fix para conservar todos los avances de UI y estabilidad.
- Al cargar el archivo Excel, el sistema mantiene el mensaje de importación exitosa y ejecuta inmediatamente el análisis sin requerir clic en "Aplicar", usando la fecha de corte disponible.
- En la pestaña "Inventario por Bodega" se eliminó la columna "Categoría" para compactar la tabla.

