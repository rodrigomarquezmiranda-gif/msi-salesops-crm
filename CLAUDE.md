# MSI SalesOps CRM — Guía de Contexto para Claude

Este archivo permite retomar el trabajo desde cualquier máquina sin perder contexto.
Rodrigo (rodrigomarquezmiranda@gmail.com) es el dueño del proyecto.

---

## Qué es esto

Aplicación web **single-page** de gestión comercial para MSI Argentina. Permite analizar ventas, stock, rotación, ranking de clientes, órdenes de compra e importaciones de los distribuidores autorizados de MSI.

La app vive en **un único archivo**: `index.html`. Toda la lógica JS, HTML y CSS está inline. El service worker está en `sw.js`.

---

## Archivos del repositorio

| Archivo | Propósito |
|---|---|
| `index.html` | La aplicación completa (todo inline) |
| `sw.js` | Service worker para PWA / caché offline |
| `favicon.png` | Ícono de la app |
| `msi-watermark.png` | Marca de agua MSI usada en exports |
| `CLAUDE.md` | Este archivo |

---

## Convención de versiones — CRÍTICO

**Formato: `DD.MM.AA.N`**

- `DD` = día (2 dígitos)
- `MM` = mes (2 dígitos)
- `AA` = año (2 dígitos)
- `N` = número secuencial de cambios del día (empieza en 1, sube con cada cambio)

Ejemplo: primer cambio del 21 de agosto de 2026 → `21.08.26.1`

**Siempre actualizar los dos lugares en el mismo edit:**

1. `index.html` → `const APP_VERSION = 'DD.MM.AA.N';` (buscar con grep)
2. `sw.js` → `const CACHE = 'salesops-DD.MM.AA.N';` (línea 3)

Si no se sincronizan, el service worker no invalida el caché y los usuarios ven la versión vieja.

---

## Distribuidores (distribuidores mayoristas de MSI)

```javascript
const DISTRIBUTORS = {
  air:         'AIR COMPUTERS',
  invid:       'INVID',
  solutionbox: 'SOLUTION BOX',
  ashir:       'ASHIR',
  elit:        'ELIT'
}
```

También existe `ETAILERS = { compragamer: 'COMPRAGAMER' }` para el canal online.

---

## Base de datos (`db`)

Todo se persiste en **Firebase Realtime Database** y se carga al iniciar. La estructura en memoria:

```javascript
let db = {
  pi: [],              // Purchase Invoices (PIs) de MSI → distribuidores
  ber: [],             // Back-end Rebates
  sor: [],             // Sales Out Reports
  sir: [],             // Stock In Reports
  pp: [],              // Price Protection
  weeklyReports: [],   // Reportes semanales de ventas/stock por distribuidor
  imports: [],         // Datos de importación por categoría
  resellers: { t1:[], t2:[], t3:[] },
  resellerTargets: {},
  resellerManual: {}
}
```

**`db.weeklyReports`** es la fuente principal de datos de rotación y clientes. Cada entrada tiene:
- `distributor` (key de DISTRIBUTORS)
- `reportType`: `'sales'` | `'stock'`
- `periodEnd`: fecha ISO string
- `rows`: array de `{ description, code, qty, avgPrice, client, cuit, date }`

**`priceList`** contiene la lista de precios LATAM con:
- `priceList.products[].qtyCtn` (col G) — unidades por carton
- `priceList.products[].qtyPlt` (col H) — unidades por pallet
- `priceList.products[].price` / `priceList.products[].priceR` (cols Q/R) — precios

---

## Tabs / secciones de la app

| Tab ID | Nombre visible | Función principal |
|---|---|---|
| `tab-dashboard` | Dashboard | KPIs generales |
| `tab-pi` | Órdenes | Purchase Invoices + botón "Armar Orden" |
| `tab-rebates` | Rebates | Back-end rebates |
| `tab-pricelist` | Lista de Precios | Lista de precios LATAM importada desde Excel |
| `tab-salesprices` | Precios de Venta | Historial de precios a los que se vendió cada modelo |
| `tab-rotation` | Rotación | Rotación de stock por distribuidor |
| `tab-trend` | Tendencias | Tendencias de ventas |
| `tab-rotcompare` | Comparador | Comparador de rotación entre distribuidores |
| `tab-clients` | Ranking de Clientes | Ranking de clientes finales por distribuidor |
| `tab-resellers-t1/t2/t3` | Resellers | Resellers por tier |
| `tab-catranking` | Categorías | Ranking por categoría |
| `tab-imports` | Importaciones | Análisis de importaciones por categoría/marca |
| `tab-etailers` | Etailers | Canal online (Compragamer) |
| `tab-semaforo` | Semáforo | Semáforo de stock vs ventas |
| `tab-share` | Market Share | Share por categoría (MSI vs competencia) |
| `tab-pbm` | PBM | Price below market analysis |
| `tab-ia` | IA | Análisis con IA |

---

## Funciones clave

### Datos de ventas
- `getSalesPriceRows()` → todas las ventas históricas de PIs subidas
- `getClientRanking(distKey, monthKey)` → agrega clientes desde `db.weeklyReports`
- `getProductComparisonGroups()` → agrupa productos por modelo key para el Comparador
- `getModelGroupComparison(modelKey, monthKey)` → stats por distribuidor para un modelo
- `extractProductModelKey(description)` → normaliza descripción a clave de modelo
- `classifyProductCategory(description)` → clasifica producto en categoría

### Clientes
- `normCuit(cuit)` → normaliza CUIT quitando guiones
- `normClientKey(name)` → normaliza nombre de cliente (mayúsculas, sin puntos)
- `distScopeList(distKey)` → lista de distribuidores según filtro ('all' o uno específico)

### UI / rendering
- `openOverlay(id)` / `closeOverlay(id)` → modales
- `positionTooltipNearElement(el)` → posiciona el tooltip compartido
- `hideClientCompositionTooltip()` / `clearTooltipHideTimer()` → sistema de tooltips
- `esc(str)` → escapa HTML
- `fmtRotDate(dateStr)` → formatea fecha para visualización
- `getMonthLabel(monthKey)` → nombre del mes para mostrar
- `getMonthKey(dateStr)` → extrae `'YYYY-MM'` de una fecha ISO
- `toast(msg)` / `_alert(msg, opts)` → notificaciones

---

## Componentes de UI reutilizables

### Tooltip compartido
```html
<div id="client-composition-tooltip">...</div>
```
Se usa en Ranking de Clientes, Comparador de Rotación, etc. Siempre:
1. `clearTooltipHideTimer()` al entrar
2. Setear `.innerHTML` y `.style.display = 'block'`
3. `positionTooltipNearElement(evt.currentTarget)`
4. `onmouseleave="hideClientCompositionTooltip()"`

### Custom dropdown (estilo PBM)
Div flotante `position:fixed` con `getBoundingClientRect()`. Usa `var(--bg)`, `var(--border2)`. No usar `<datalist>` nativo (lento y blanco en dark mode).

---

## Armar Orden (modal en tab Órdenes)

Función: `openOrderBuilder()` → modal `#modal-order-builder`

- Distribuidor: `<select id="ob-client">` con los 5 distys
- Modelo: input con dropdown custom (`#ob-suggest-drop`) buscando en `priceList.products`
- Auto-precio: busca precio mínimo histórico con `_obPriceHistMap()`
- Hint packaging: muestra `qtyCtn` / `qtyPlt` en cursiva bajo el modelo
- Export: `exportOrderToExcel()` — ExcelJS con estilo navy/orange
- Mail: `sendOrderByEmail()` → preview en `#modal-order-email` → `_obCopyEmailHtml()` copia HTML enriquecido al clipboard

---

## Comparador de Rotación (tab-rotcompare)

Estado global:
```javascript
let currentRotCompareRows = [];   // rows del distribuidor detail actual
let _currentRotModelKey = null;   // modelo seleccionado actualmente
let _rotExpandedDist = null;      // distribuidor expandido (click en fila)
let _rotExpandedSort = 'price_asc'; // orden del expand: 'price_asc'|'price_desc'|'date'
let _rotExpandSeeAll = false;     // ver todos vs primeros 20
let rotcomparePeriod = 'all';     // período seleccionado
```

Funciones de interacción:
- `showMinPriceClientTooltip(evt, distKey)` → hover en precio mín → muestra últimas 5 ventas más baratas
- `toggleRotDistExpand(distKey)` → click en fila distribuidor → expande/colapsa lista de ventas
- `setRotExpandSort(s)` / `setRotExpandSeeAll(v)` → controles del panel expandido
- `_getRotDistSales(distKey, modelKey, monthKey)` → obtiene todas las ventas para expand + tooltip

---

## Ranking de Clientes (tab-clients)

Estado:
```javascript
let selectedClientKey = null;
let _clientDetailSearch = '';
let _cdDistKey = '', _cdMonthKey = '', _cdClientKey = ''; // parámetros del render actual
```

Funciones:
- `renderClientsTab()` → render principal
- `renderClientDetail(distKey, monthKey, clientKey)` → render del panel derecho (frame completo)
- `_filterClientDetailTable()` → filtrado rápido del buscador (solo toca `tbody` + contador, no rebuilda el DOM)
- `showClientProductTooltip(evt, clientKey, descNorm, distKey)` → hover en producto → últimas 5 compras + min/max/avg precio

El buscador en el detalle de cliente usa `oninput="_clientDetailSearch=this.value;_filterClientDetailTable()"` para evitar re-render completo en cada tecla.

---

## Categorías de producto

```javascript
// Keys: motherboard, monitor, vga, gabinete, psu, minipc, ssd, cooling, accesorios, other
const CATEGORY_KEYWORDS = [...] // keywords por categoría, orden importa (más específico primero)
const CATEGORY_COLOR = { motherboard:'#88C0D0', monitor:'#81A1C1', vga:'#A3BE8C', ... }
```

---

## Dependencias externas (CDN, no instalar)

- **Firebase** 9.23.0 — auth + realtime database
- **ExcelJS** — exports Excel estilizados (navy/orange branding)
- **SheetJS (XLSX)** — lectura de archivos Excel subidos
- **Tabler Icons** (`ti ti-*`) — iconografía
- **jsPDF + AutoTable** — exports PDF

---

## Colores de marca

| Variable CSS | Valor típico | Uso |
|---|---|---|
| `var(--brand)` | `#FF6600` | Naranja MSI (bordes, highlights) |
| `var(--brand-text)` | naranja más suave | Texto en naranja |
| `var(--bg)` / `var(--bg2)` / `var(--bg3)` | grises oscuros | Fondos |
| `var(--text)` / `var(--text2)` / `var(--text3)` | blancos/grises | Texto |
| `var(--border)` / `var(--border2)` | grises | Bordes |
| `var(--green-text)` | verde | Precio mínimo, indicadores positivos |
| `var(--blue-text)` | azul | Cobertura de semanas |

Para Excel exports: NAVY `'FF1F2A44'`, ORANGE `'FFFF6600'`, WHITE `'FFFFFFFF'`, GRAY `'FFF5F5F5'`

---

## Reglas de trabajo

1. **Un solo archivo**: todos los cambios van en `index.html`. No crear archivos JS/CSS separados.
2. **Versión**: actualizar `APP_VERSION` en `index.html` Y `CACHE` en `sw.js` en cada cambio. Formato `DD.MM.AA.N`.
3. **No usar `<datalist>` nativo** para búsquedas: usar dropdown custom con `position:fixed`.
4. **No re-renderizar DOM completo en cada keystroke**: separar frame render de filter render (ver `_filterClientDetailTable`).
5. **Precios**: siempre tomar valor absoluto (`Math.abs`) — notas de crédito pueden tener negativos.
6. **Git**: el repo está en GitHub. Archivos auto-generados (`stock_changelog.json`, `stock_diff_latest.json`) pueden tener merge conflicts — resolver quedándose con "theirs" (la versión remota).
