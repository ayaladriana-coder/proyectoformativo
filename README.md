# proyectoformativo
# Control Mensual Centro — Tablero + Regresión lineal

Página **HTML + CSS + JS** que corre 100 % en el navegador (sin servidor, sin Python).
Tiene dos pestañas:

| Pestaña | Qué hace |
|---|---|
| **Tablero** | El tablero del mes que ya existía (KPIs y 4 gráficos). No cambió su lógica. |
| **Regresión** | Carga los datos, arma un modelo de regresión lineal (simple o múltiple), muestra métricas, gráficos, diagnóstico y permite predecir. |

> Abrí `index.html` haciendo doble clic. No hace falta instalar nada: PapaParse, Plotly y
> SheetJS se cargan por CDN. Si la máquina está sin internet, cargá esos tres archivos
> localmente y cambiá los `src` del `<head>`.

---

## 1. Cómo usar la pestaña Regresión (5 pasos en una pantalla)

1. **Cargar** — arrastrá el archivo a la zona punteada (o hacé clic). Acepta
   - **CSV transaccional**: `fecha, dia_semana, seccion, item, cantidad, precio, monto` (el `control_mensual_sept.csv`).
   - **CSV / hoja matriz diaria**: un concepto por fila y un día por columna (como salen `ventas`, `caja`, `gastos`, `insumos`).
   - **`.xlsx` / `.xlsm` directo**: se lee con SheetJS; si el libro tiene varias hojas, elegís la hoja en el selector. Las macros se ignoran: solo se leen valores (y fórmulas ya calculadas).
2. **Unidad + variables** — elegí la unidad de análisis, la hoja (si aplica), la **Y** (una sola) y una o más **X**. El tipo de cada X se puede alternar (numérica ↔ categórica) con el enlace de tipo. Las categóricas generan variables dummy.
3. **Ajustar** — marcá/desmarcá intercepto, elegí el `k` de la validación cruzada y tocá **Ajustar modelo**. Con 1 X se comporta como regresión simple; con 2 o más, como múltiple.
4. **Ver** — ecuación, coeficientes (con error estándar, t, p, IC 95 %, β estandarizado), métricas de ajuste y de error, diagnóstico (VIF, Durbin-Watson, outliers de Cook) y los 7 gráficos.
5. **Predecir** — completá los valores de las X y obtené **Ŷ** con IC 95 % para la media y para una observación nueva. Podés cargar varios **escenarios** y exportarlos a CSV.

### Unidades de análisis

| Unidad | Filas | Y típica | Predictoras típicas |
|---|---|---|---|
| **Día** (por defecto) | 28 | `monto_total` | `cantidad_grandes`, `cantidad_medianos`, `ticket_promedio`, `dia_mes`, `fin_de_semana`, `dia_semana`… |
| **Ítem-día** | 330 | `monto` de la línea | `cantidad`, `precio`, `seccion`, `item`, `dia_semana` |

### Variables que la app construye sola (nivel día)

`monto_total`, `unidades_total`, `cantidad_grandes`, `cantidad_medianos`, `monto_grandes`,
`monto_medianos`, `monto_toppings`, `monto_pasteleria`, `monto_bubblebath`, `monto_merchandise`,
`monto_extras`, `pedidos_ya`, `unidades_promo`, `ticket_promedio`, `precio_promedio_vaso`,
`items_distintos`, `dia_mes` (tendencia), `dia_semana` (categórica) y `fin_de_semana` (0/1).

Si cargás una hoja matriz (por ejemplo `caja`), cada concepto aparece como una variable más
(`efectivo`, `qr_neto`, `tarjeta_neto`, `total_gastos`, …) y se suman a la matriz diaria.

---

## 2. Cómo se calcula (motor estadístico)

Todo está en `js/estadistica.js`, en JS puro (sin librerías de cálculo):

- **Mínimos cuadrados ordinarios resueltos por descomposición QR (Householder)** en vez de invertir `X'X` a mano → más estable numéricamente.
- **Ajuste**: R², R² ajustado, n, k, grados de libertad, σ (error estándar de la estimación), F global con p-valor.
- **Error**: MAE, RMSE, MAPE, error medio (sesgo) y **validación cruzada k-fold** (R², RMSE y MAE fuera de muestra).
- **Por coeficiente**: β, error estándar, t, p-valor, IC 95 % y β estandarizado.
- **Diagnóstico**: VIF y tolerancia, Durbin-Watson, apalancamiento, residuo estudentizado, distancia de Cook, outliers, asimetría y curtosis.
- **Distribuciones** t de Student y F implementadas con la función beta incompleta (Lanczos + fracción continua), igual que en libros de recetas numéricas.
- **Predicción**: `Ŷ`, IC 95 % para la media y para una observación nueva (`σ²·(1 + x₀'(X'X)⁻¹x₀)`), con aviso de **extrapolación** cuando un valor cae fuera del rango observado.

### Controles de realismo

- Con menos de ~10 observaciones por predictora la app **avisa** (n = 28 en el nivel día → conviene usar 2–3 X como máximo).
- Si un VIF supera 10, lo marca como colineal.
- Las variables de caja (efectivo/QR/tarjeta) son **consecuencia** de las ventas, no causa: se señalan como tales.
- Priorizá siempre el **R² de validación cruzada**, no el R² del ajuste.

---

## 3. Gráficos

1. Dispersión X vs Y + recta ajustada con bandas de confianza y predicción 95 % (regresión simple; gráfico representativo).
2. Y real vs Ŷ con recta de 45° y R² anotado (múltiple).
3. Residuos vs valores ajustados.
4. QQ-plot de residuos.
5. Coeficientes estandarizados ± IC 95 % (importancia).
6. Regresión parcial (added-variable) por cada X, seleccionable.
7. Serie temporal: Y real vs predicho día a día.

---

## 4. Archivos

```
index.html        Pestañas [Tablero][Regresión]. La pestaña Tablero conserva su lógica original.
css/app.css       Barra de pestañas y estilos de la app de regresión (paleta azul actual).
js/estadistica.js Motor OLS + métricas + distribuciones + VIF + validación cruzada + predicción.
js/datos.js       Carga CSV/XLSX, reconocimiento de variables y pivote día / ítem-día.
js/graficos.js    Los 7 gráficos con Plotly.
js/app.js         Estado de la UI, pestañas, selección Y/X, resultados, predicción y escenarios.

control_mensual_sept.csv   Datos de septiembre 2026 (formato transaccional).
Consolidados.bas           Módulo VBA de consolidados (pendiente: macro de export diario, ver §6).
COMO_USAR.md               Guía de las macros y del libro de Excel.
PLAN_REGRESION.md          Plan/decisión del aplicativo.

_val/                      Suite de verificación opcional (Node + Chrome) — ver §5.
```

Sin módulos ES: los `js/*.js` se cargan con `<script src>` clásico y exponen objetos globales
(`window.Estadistica`, `window.Datos`, `window.Graficos`), porque `type="module"` falla bajo `file://`.

---

## 5. Verificación

La suite de `_val/` corre con Node y Chrome y se apoya en `numpy`/`scipy`/`statsmodels` como referencia.

```bash
cd _val
node validar.js && python validar.py     # motor OLS vs statsmodels
python -c "..."                          # comparación numérica (incluida abajo)
node validar_datos.js                     # criterios de aceptación con los JS reales
node smoke.js                             # end-to-end en Chrome headless (drag&drop, ajuste, gráficos)
```

Resultados obtenidos:

- **Motor vs statsmodels**: diferencia máxima absoluta **5,8e-13** en β, EE, t, p, R², σ, F, IC — muy por debajo del criterio `1e-6`.
- **Carga de `control_mensual_sept.csv`**: detecta formato transaccional, **28 días**, 330 filas, 37 ítems y **venta total $30.016,8**.
- **Modelo simple** (`monto_total ~ cantidad_grandes`): R² = 0,4310. **Modelo múltiple** (`+ cantidad_medianos + ticket_promedio`): R² = 0,9605, R² ajustado = 0,9556, F = 194,62 (p = 5,7e-17).
- **Predicción** para `(5, 8, 30)`: Ŷ = 415,743 con IC 95 % de la media `[294,2; 537,3]` y de una observación nueva `[226,1; 605,4]`, idéntico a statsmodels.
- **Smoke test en Chrome**: carga por arrastrar y soltar, cambio de pestañas, modelo simple y múltiple, los 7 gráficos, predicción, aviso de extrapolación, escenarios y export — **sin excepciones ni errores de consola**.
- **VIF**: coincide con `statsmodels` (máx. diferencia 0,0004) y detecta colinealidad (VIF > 10) cuando existe.

---

## 6. Pendiente (Fase 5 del plan)

La macro **`ExportarDatosAnalisis`** en `Consolidados.bas`, que escribe `datos/datos_diarios.csv`
con fecha, ventas por sección, caja, gastos e insumos (una fila por día), **todavía no está incluida**.
Mientras tanto el **CSV transaccional** ya cubre ventas por ítem/sección; caja, gastos e insumos se
pueden cargar igual arrastrando directamente el `.xlsm` y eligiendo la hoja (`caja`, `gastos`, `insumos`).

---

## 7. Limitaciones a tener presentes

- **n = 28 días.** Con muchas X el R² sube artificialmente. Regla práctica: máximo 2–3 predictores.
- La multicolinealidad es esperable (grandes y medianos comparten estructura con las unidades totales).
- Septiembre tiene promociones puntuales (Promo 50 %, stickers, Mediano Promo premium) que actúan como outliers; la distancia de Cook los señala.
- Correlación no es causalidad: los resultados describen el mes, no son una ley.
