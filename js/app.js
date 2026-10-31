/**
 * js/app.js
 * Orquestador Principal — Suite de Inteligencia Comercial & Pricing INTERFRUTD
 * Cubre:
 *  1. Gestión y actualización de Listas de Precios, Costos de Reposición y Benchmark (Supermercados & CENADA)
 *  2. Elaboración de Cotizaciones Comerciales B2B para Clientes (con simulador de margen y conversión a OC)
 *  3. Seguimiento de Órdenes de Compra (OC), Ventas, Inventarios, Cobertura y Clasificación ABC
 *  4. Query Builder Comercial Visual + Exportación Dinámica a Excel Avanzado (.xlsx)
 */

import {
    INTERFRUTD_CATALOG,
    CENADA_BENCHMARK,
    B2B_CLIENTS,
    INITIAL_QUOTES,
    INITIAL_PURCHASE_ORDERS
} from './data/interfrutdData.js';
import {
    formatCRC,
    formatNum,
    formatPct,
    escapeHtml,
    showToast,
    exportRowsToExcel,
    exportMultiSheetExcel
} from './utils.js';

// ============================================================================
// ESTADO REACTIVO DE LA APLICACIÓN
// ============================================================================
const appState = {
    activeTab: 'pricing',
    searchQuery: '',
    globalCategory: 'ALL',
    globalAbc: 'ALL',

    // Módulo 1: Pricing & Listas
    pricingView: 'catalog', // 'catalog' | 'cenada'
    activePriceListChannel: 'ALL',
    targetMarginSimPct: 30,
    pricingDiagFilter: 'ALL',
    catalog: [],
    cenadaList: [...CENADA_BENCHMARK],

    // Módulo 2: Cotizaciones B2B
    clients: [...B2B_CLIENTS],
    quotes: [],
    activeQuoteDraft: {
        id: 'COT-2026-0145',
        fecha: new Date().toISOString().slice(0, 10),
        vigenciaHasta: '2026-11-08',
        clienteId: 'CLI-001',
        listaPrecio: 'HORECA',
        notas: 'Precios puestos en andén del cliente. Producto fresco seleccionado bajo estándar Interfrutd El Fruto Perfecto.',
        items: []
    },

    // Módulo 3: Órdenes de Compra, Ventas e Inventarios
    purchaseOrders: [],
    ocFilterStatus: 'ALL',
    inventoryFilter: 'ALL',

    // Módulo 4: Query Builder Comercial
    qbPreset: 'criticos_abc_a',
    qbDataset: 'MASTER_JOIN',
    qbFilterField: 'margenReposicionPct',
    qbFilterOp: '<',
    qbFilterVal: '22',
    qbGroupBy: 'NONE',
    qbOrderBy: 'ventasMensualesCrc_DESC',
    lastQueryExportRows: []
};

let chartCatInstance = null;
let chartAbcInstance = null;

// ============================================================================
// INICIALIZACIÓN Y RESTAURACIÓN DE ESTADO LOCAL
// ============================================================================
function initStateData() {
    // Clonar catálogo base y aplicar precios personalizados si existen en localStorage
    let savedPriceOverrides = {};
    try {
        const raw = localStorage.getItem('interfrutd_price_overrides');
        if (raw) savedPriceOverrides = JSON.parse(raw) || {};
    } catch (e) {}

    appState.catalog = INTERFRUTD_CATALOG.map(item => {
        const copy = { ...item };
        if (savedPriceOverrides[copy.sku] !== undefined) {
            copy.precioDetalle = Number(savedPriceOverrides[copy.sku]);
        }
        recalcSkuMetrics(copy);
        return copy;
    });

    // Cotizaciones
    try {
        const savedQ = localStorage.getItem('interfrutd_quotes');
        appState.quotes = savedQ ? JSON.parse(savedQ) : structuredClone(INITIAL_QUOTES);
    } catch (e) {
        appState.quotes = structuredClone(INITIAL_QUOTES);
    }

    // Órdenes de compra
    try {
        const savedOC = localStorage.getItem('interfrutd_purchase_orders');
        appState.purchaseOrders = savedOC ? JSON.parse(savedOC) : structuredClone(INITIAL_PURCHASE_ORDERS);
    } catch (e) {
        appState.purchaseOrders = structuredClone(INITIAL_PURCHASE_ORDERS);
    }

    // Inicializar borrador de cotización con la primera cotización si está vacío
    if (appState.quotes.length > 0 && appState.activeQuoteDraft.items.length === 0) {
        loadQuoteIntoBuilder(appState.quotes[0].id, false);
    }
}

function recalcSkuMetrics(item) {
    const cost = item.costoReposicion > 0 ? item.costoReposicion : item.costoActual;
    const price = item.precioDetalle > 0 ? item.precioDetalle : 1;
    item.margenReposicionPct = Math.round(((price - cost) / price) * 1000) / 10;
    item.margenActualPct = Math.round(((price - item.costoActual) / price) * 1000) / 10;
    const targetM = appState.targetMarginSimPct / 100;
    item.precioSugeridoMeta = Math.max(15, Math.round((cost / (1 - targetM)) / 5) * 5);

    if (item.margenReposicionPct < 16) item.estadoPrecio = 'MARGEN_CRITICO';
    else if (item.varCostoPct >= 6) item.estadoPrecio = 'COSTO_SUBIO';
    else if (item.precioCompetenciaMin > 0 && price > item.precioCompetenciaMin * 1.15) item.estadoPrecio = 'CARO_VS_COMPETENCIA';
    else if (item.precioCompetenciaMin > 0 && price < item.precioCompetenciaMin * 0.85 && item.margenReposicionPct < appState.targetMarginSimPct) item.estadoPrecio = 'OPORTUNIDAD_PRECIO';
    else item.estadoPrecio = 'OPTIMO';
}

function getFilteredCatalog() {
    const q = appState.searchQuery.trim().toLowerCase();
    return appState.catalog.filter(item => {
        if (appState.globalCategory !== 'ALL' && item.categoria !== appState.globalCategory) return false;
        if (appState.globalAbc !== 'ALL' && item.abcClass !== appState.globalAbc) return false;
        if (q) {
            const match = item.sku.toLowerCase().includes(q) ||
                item.name.toLowerCase().includes(q) ||
                item.categoria.toLowerCase().includes(q) ||
                (item.especie && item.especie.toLowerCase().includes(q));
            if (!match) return false;
        }
        return true;
    });
}

// ============================================================================
// MÓDULO 1: PRICING, COSTOS, LISTAS DE PRECIOS & BENCHMARK
// ============================================================================
function renderPricingModule() {
    const baseList = getFilteredCatalog();
    baseList.forEach(recalcSkuMetrics);

    // KPIs de Pricing
    const totalSkus = baseList.length;
    const totalSales = baseList.reduce((acc, x) => acc + x.ventasMensualesCrc, 0);
    const totalProfit = baseList.reduce((acc, x) => acc + (x.demandaMensualUnd * (x.precioDetalle - x.costoReposicion)), 0);
    const totalRevRetail = baseList.reduce((acc, x) => acc + (x.demandaMensualUnd * x.precioDetalle), 0);
    const weightedMarginPct = totalRevRetail > 0 ? (totalProfit / totalRevRetail) * 100 : 0;

    const criticalCount = baseList.filter(x => x.margenReposicionPct < 16).length;
    const costUpCount = baseList.filter(x => x.varCostoPct >= 6).length;
    const withBench = baseList.filter(x => x.precioCompetenciaMin > 0);
    const cheaperThanSuperCount = withBench.filter(x => x.precioDetalle <= x.precioCompetenciaMin).length;
    const competitiveIndexPct = withBench.length > 0 ? (cheaperThanSuperCount / withBench.length) * 100 : 100;

    const kpiContainer = document.getElementById('pricingKpiGrid');
    if (kpiContainer) {
        kpiContainer.innerHTML = `
            <div class="kpi-card">
                <div class="kpi-header"><span>SKUs Activos en Lista</span><i class="fa-solid fa-boxes-stacked"></i></div>
                <div class="kpi-value">${formatNum(totalSkus)}</div>
                <div class="kpi-sub">
                    <span class="badge badge-success">${baseList.filter(x => x.abcClass === 'A').length} Clase A</span>
                    <span>4 Listas de Precios activas</span>
                </div>
            </div>
            <div class="kpi-card ${weightedMarginPct < 24 ? 'kpi-warning' : ''}">
                <div class="kpi-header"><span>Margen Bruto Ponderado</span><i class="fa-solid fa-percent"></i></div>
                <div class="kpi-value">${formatPct(weightedMarginPct, 1)}</div>
                <div class="kpi-sub">
                    <span>Meta simulada: <strong>${appState.targetMarginSimPct}%</strong></span>
                    <span>· Utilidad: ${formatCRC(totalProfit)}/mes</span>
                </div>
            </div>
            <div class="kpi-card ${criticalCount > 0 ? 'kpi-danger' : ''}">
                <div class="kpi-header"><span>Alertas de Costo & Margen</span><i class="fa-solid fa-triangle-exclamation"></i></div>
                <div class="kpi-value">${criticalCount} <span style="font-size:14px; font-weight:600; color:var(--text-muted);">SKUs &lt;16%</span></div>
                <div class="kpi-sub">
                    <span class="badge badge-warning">${costUpCount} subieron de costo (&ge;6%)</span>
                </div>
            </div>
            <div class="kpi-card kpi-info">
                <div class="kpi-header"><span>Índice Competitivo vs Superm.</span><i class="fa-solid fa-scale-balanced"></i></div>
                <div class="kpi-value">${formatPct(competitiveIndexPct, 0)}</div>
                <div class="kpi-sub">
                    <span>${cheaperThanSuperCount} de ${withBench.length} SKUs más competitivos que AutoMercado/Walmart</span>
                </div>
            </div>
        `;
    }

    const badgePricing = document.getElementById('badgeCountPricing');
    if (badgePricing) badgePricing.textContent = String(totalSkus);

    // Si estamos en vista PIMA CENADA
    const catalogCont = document.getElementById('pricingCatalogContainer');
    const cenadaCont = document.getElementById('cenadaComparisonContainer');
    const controlsBar = document.getElementById('pricingControlsBar');
    if (appState.pricingView === 'cenada') {
        if (catalogCont) catalogCont.style.display = 'none';
        if (controlsBar) controlsBar.style.display = 'none';
        if (cenadaCont) cenadaCont.style.display = 'block';
        renderCenadaComparisonTable();
        return;
    } else {
        if (catalogCont) catalogCont.style.display = 'block';
        if (controlsBar) controlsBar.style.display = 'flex';
        if (cenadaCont) cenadaCont.style.display = 'none';
    }

    // Filtrar por chip de diagnóstico
    const rows = baseList.filter(item => {
        if (appState.pricingDiagFilter === 'MARGEN_CRITICO') return item.margenReposicionPct < 16;
        if (appState.pricingDiagFilter === 'COSTO_SUBIO') return item.varCostoPct >= 6;
        if (appState.pricingDiagFilter === 'OPORTUNIDAD_PRECIO') return item.estadoPrecio === 'OPORTUNIDAD_PRECIO';
        if (appState.pricingDiagFilter === 'CON_BENCHMARK') return item.precioCompetenciaMin > 0;
        return true;
    });

    const tbody = document.getElementById('pricingTableBody');
    if (!tbody) return;

    const ch = appState.activePriceListChannel;

    tbody.innerHTML = rows.map(item => {
        const varBadge = item.varCostoPct > 0
            ? `<span class="badge ${item.varCostoPct >= 6 ? 'badge-danger' : 'badge-warning'}">+${formatPct(item.varCostoPct, 1)}</span>`
            : (item.varCostoPct < 0
                ? `<span class="badge badge-success">${formatPct(item.varCostoPct, 1)}</span>`
                : `<span class="badge badge-neutral">0.0%</span>`);

        let marginBadgeClass = 'badge-success';
        if (item.margenReposicionPct < 16) marginBadgeClass = 'badge-danger';
        else if (item.margenReposicionPct < 24) marginBadgeClass = 'badge-warning';

        let diagHtml = `<span class="badge badge-success"><i class="fa-solid fa-check"></i> Óptimo</span>`;
        if (item.estadoPrecio === 'MARGEN_CRITICO') {
            diagHtml = `<span class="badge badge-danger"><i class="fa-solid fa-bolt"></i> Margen Crítico</span>`;
        } else if (item.estadoPrecio === 'COSTO_SUBIO') {
            diagHtml = `<span class="badge badge-warning"><i class="fa-solid fa-arrow-trend-up"></i> Costo Subió</span>`;
        } else if (item.estadoPrecio === 'OPORTUNIDAD_PRECIO') {
            diagHtml = `<span class="badge badge-info"><i class="fa-solid fa-circle-up"></i> Oportunidad Alza</span>`;
        } else if (item.estadoPrecio === 'CARO_VS_COMPETENCIA') {
            diagHtml = `<span class="badge badge-warning"><i class="fa-solid fa-scale-unbalanced"></i> Sobre Competencia</span>`;
        }

        const benchHtml = item.precioCompetenciaMin > 0
            ? `<div title="AutoMercado: ${formatCRC(item.precioAutoMercado)} | Walmart: ${formatCRC(item.precioWalmart)} | MásxMenos: ${formatCRC(item.precioMasxMenos)}">
                 <strong>${formatCRC(item.precioCompetenciaMin)}</strong>
               </div>`
            : `<span style="color:var(--text-muted);">—</span>`;

        const highlightStyle = (colName) => (ch === colName ? 'background: rgba(77,184,72,0.13); font-weight:800;' : '');

        return `
            <tr>
                <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(item.sku)}</code></td>
                <td>
                    <div style="font-weight:700;">${escapeHtml(item.name)}</div>
                    <div style="font-size:11px; color:var(--text-muted);">Empaque: ${item.unidad} (x${item.factorEmpaque}) · Stock: ${formatNum(item.stockTotal)}</div>
                </td>
                <td><span class="badge badge-neutral">${escapeHtml(item.categoria)}</span></td>
                <td class="center-col"><span class="abc-badge abc-${item.abcClass}">${item.abcClass}</span></td>
                <td class="num-col">${formatCRC(item.costoActual)}</td>
                <td class="num-col" style="font-weight:700;">${formatCRC(item.costoReposicion)}</td>
                <td class="num-col">${varBadge}</td>
                <td class="num-col" style="${highlightStyle('MAYORISTA')}">${formatCRC(item.precioMayorista)}</td>
                <td class="num-col" style="${highlightStyle('SUPERMERCADO')}">${formatCRC(item.precioSupermercado)}</td>
                <td class="num-col" style="${highlightStyle('HORECA')}">${formatCRC(item.precioHoreca)}</td>
                <td class="num-col" style="${highlightStyle('DETALLE')}">
                    <input type="number" class="inline-price-input" data-sku="${escapeHtml(item.sku)}" value="${item.precioDetalle}" min="5" step="5">
                </td>
                <td class="num-col"><span class="badge ${marginBadgeClass}">${formatPct(item.margenReposicionPct, 1)}</span></td>
                <td class="num-col" style="font-weight:700; color:var(--brand-lime);">${formatCRC(item.precioSugeridoMeta)}</td>
                <td class="num-col">${benchHtml}</td>
                <td>${diagHtml}</td>
                <td class="center-col">
                    <button type="button" class="btn btn-secondary btn-sm btn-quick-quote" data-sku="${escapeHtml(item.sku)}" title="Enviar a Cotización B2B">
                        <i class="fa-solid fa-file-circle-plus"></i>
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

function renderCenadaComparisonTable() {
    const tbody = document.getElementById('cenadaTableBody');
    if (!tbody) return;
    const q = appState.searchQuery.trim().toLowerCase();

    const filtered = appState.cenadaList.filter(row => {
        if (!q) return true;
        return String(row.frumusaCod || '').toLowerCase().includes(q) ||
            String(row.frumusaName || '').toLowerCase().includes(q) ||
            String(row.cenadaName || '').toLowerCase().includes(q);
    });

    tbody.innerHTML = filtered.map(r => {
        const diff = Number(r.diffColones) || 0;
        const pct = Number(r.pctDiffModa) || 0;
        let badge = `<span class="badge badge-info">Alineado al Mercado</span>`;
        if (pct <= -5) badge = `<span class="badge badge-success"><i class="fa-solid fa-trophy"></i> Ventaja Competitiva (${formatPct(pct, 1)})</span>`;
        else if (pct >= 8) badge = `<span class="badge badge-warning"><i class="fa-solid fa-triangle-exclamation"></i> Sobre Moda CENADA (+${formatPct(pct, 1)})</span>`;

        return `
            <tr>
                <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(r.frumusaCod)}</code></td>
                <td style="font-weight:700;">${escapeHtml(r.frumusaName)}</td>
                <td><span class="badge badge-neutral">${escapeHtml(r.frumusaUnit)}</span></td>
                <td class="num-col" style="font-weight:800;">${formatCRC(r.frumusaPrecioIva)}</td>
                <td>${escapeHtml(r.cenadaName)}</td>
                <td class="num-col">${formatCRC(r.cenadaMin)}</td>
                <td class="num-col" style="font-weight:700;">${formatCRC(r.cenadaModa)}</td>
                <td class="num-col">${formatCRC(r.cenadaAvg)}</td>
                <td class="num-col" style="color:${diff <= 0 ? 'var(--success)' : 'var(--warning)'}; font-weight:700;">
                    ${diff > 0 ? '+' : ''}${formatCRC(diff)}
                </td>
                <td class="num-col">${pct > 0 ? '+' : ''}${formatPct(pct, 1)}</td>
                <td>${badge}</td>
            </tr>
        `;
    }).join('');
}

// ============================================================================
// MÓDULO 2: COTIZADOR COMERCIAL B2B PARA CLIENTES
// ============================================================================
function populateQuoteSelectors() {
    const clientSel = document.getElementById('quoteClientSelect');
    if (clientSel && clientSel.options.length === 0) {
        clientSel.innerHTML = appState.clients.map(c =>
            `<option value="${escapeHtml(c.id)}">${escapeHtml(c.nombre)} (${escapeHtml(c.segmento)})</option>`
        ).join('');
    }

    const skuSel = document.getElementById('quoteAddSkuSelect');
    if (skuSel && skuSel.options.length === 0) {
        const sorted = [...appState.catalog].sort((a, b) => a.name.localeCompare(b.name));
        skuSel.innerHTML = sorted.map(item =>
            `<option value="${escapeHtml(item.sku)}">[${escapeHtml(item.sku)}] ${escapeHtml(item.name)} — HORECA: ${formatCRC(item.precioHoreca)} | Stock: ${formatNum(item.stockTotal)}</option>`
        ).join('');
    }
}

function getBasePriceByList(skuObj, listType) {
    if (!skuObj) return 0;
    if (listType === 'MAYORISTA') return skuObj.precioMayorista;
    if (listType === 'SUPERMERCADO') return skuObj.precioSupermercado;
    if (listType === 'DETALLE') return skuObj.precioDetalle;
    return skuObj.precioHoreca;
}

function renderQuotesModule() {
    populateQuoteSelectors();

    const draft = appState.activeQuoteDraft;
    const titleEl = document.getElementById('quoteBuilderTitle');
    if (titleEl) titleEl.textContent = `Constructor de Cotización Comercial (${draft.id})`;

    const clientSel = document.getElementById('quoteClientSelect');
    if (clientSel) clientSel.value = draft.clienteId;

    const listSel = document.getElementById('quotePriceListSelect');
    if (listSel) listSel.value = draft.listaPrecio;

    const validIn = document.getElementById('quoteValidUntil');
    if (validIn) validIn.value = draft.vigenciaHasta;

    const notesIn = document.getElementById('quoteNotesInput');
    if (notesIn) notesIn.value = draft.notas || '';

    // Calcular totales de la cotización en edición
    let draftRevenue = 0;
    let draftCost = 0;
    const tbody = document.getElementById('quoteLinesTableBody');

    if (tbody) {
        tbody.innerHTML = draft.items.map((line, idx) => {
            const catItem = appState.catalog.find(x => x.sku === line.sku);
            const stockDisp = catItem ? catItem.stockTotal : 100;
            const precioFinal = Math.round(line.precioLista * (1 - (Number(line.descuentoPct) || 0) / 100));
            line.precioFinal = precioFinal;
            const subtotal = line.cantidad * precioFinal;
            const subtotalCost = line.cantidad * line.costoUnit;
            draftRevenue += subtotal;
            draftCost += subtotalCost;
            const lineMarginPct = precioFinal > 0 ? ((precioFinal - line.costoUnit) / precioFinal) * 100 : 0;

            const stockBadge = line.cantidad > stockDisp
                ? `<span class="badge badge-danger" title="Supera stock inmediato">${formatNum(stockDisp)} ⚠️</span>`
                : `<span class="badge badge-success">${formatNum(stockDisp)}</span>`;

            const marginClass = lineMarginPct < 16 ? 'badge-danger' : (lineMarginPct < 22 ? 'badge-warning' : 'badge-success');

            return `
                <tr>
                    <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(line.sku)}</code></td>
                    <td style="font-weight:600;">${escapeHtml(line.name)}</td>
                    <td class="center-col">${escapeHtml(line.unidad)}</td>
                    <td class="num-col">${stockBadge}</td>
                    <td class="num-col">
                        <input type="number" class="inline-price-input quote-line-qty" data-idx="${idx}" value="${line.cantidad}" min="1" style="width:72px;">
                    </td>
                    <td class="num-col">${formatCRC(line.costoUnit)}</td>
                    <td class="num-col">${formatCRC(line.precioLista)}</td>
                    <td class="num-col">
                        <input type="number" class="inline-price-input quote-line-disc" data-idx="${idx}" value="${line.descuentoPct}" min="0" max="35" step="0.5" style="width:64px;">
                    </td>
                    <td class="num-col" style="font-weight:800;">${formatCRC(precioFinal)}</td>
                    <td class="num-col" style="font-weight:700;">${formatCRC(subtotal)}</td>
                    <td class="num-col"><span class="badge ${marginClass}">${formatPct(lineMarginPct, 1)}</span></td>
                    <td class="center-col">
                        <button type="button" class="btn btn-danger btn-sm btn-remove-qline" data-idx="${idx}" title="Eliminar línea">
                            <i class="fa-solid fa-trash"></i>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');
    }

    const draftProfit = draftRevenue - draftCost;
    const draftMarginPct = draftRevenue > 0 ? (draftProfit / draftRevenue) * 100 : 0;

    let approvalBadge = `<span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> APROBACIÓN AUTOMÁTICA (&ge;22%)</span>`;
    if (draftMarginPct < 16) {
        approvalBadge = `<span class="badge badge-danger"><i class="fa-solid fa-hand"></i> MARGEN CRÍTICO (&lt;16% · Requiere Ajuste)</span>`;
    } else if (draftMarginPct < 22) {
        approvalBadge = `<span class="badge badge-warning"><i class="fa-solid fa-user-shield"></i> REQUIERE APROBACIÓN JEFATURA (16%-21.9%)</span>`;
    }

    const summaryBar = document.getElementById('quoteFinancialSummaryBar');
    if (summaryBar) {
        summaryBar.innerHTML = `
            <div style="display:flex; align-items:center; gap:14px; flex-wrap:wrap;">
                <span>Líneas: <strong>${draft.items.length}</strong></span>
                <span>Costo Total: <strong>${formatCRC(draftCost)}</strong></span>
                <span>Total Cotizado: <strong style="font-size:15px; color:var(--brand-lime);">${formatCRC(draftRevenue)}</strong></span>
                <span>Utilidad Bruta: <strong style="color:var(--success);">${formatCRC(draftProfit)} (${formatPct(draftMarginPct, 1)})</strong></span>
            </div>
            <div>${approvalBadge}</div>
        `;
    }

    // KPIs de Cartera de Cotizaciones
    let totalPortfolioRev = 0;
    let totalPortfolioProfit = 0;
    appState.quotes.forEach(q => {
        q.items.forEach(it => {
            const p = it.precioFinal || Math.round(it.precioLista * (1 - (it.descuentoPct || 0) / 100));
            totalPortfolioRev += it.cantidad * p;
            totalPortfolioProfit += it.cantidad * (p - it.costoUnit);
        });
    });
    const avgPortfolioMargin = totalPortfolioRev > 0 ? (totalPortfolioProfit / totalPortfolioRev) * 100 : 0;

    const quotesKpi = document.getElementById('quotesKpiGrid');
    if (quotesKpi) {
        quotesKpi.innerHTML = `
            <div class="kpi-card">
                <div class="kpi-header"><span>Cotizaciones en Cartera</span><i class="fa-solid fa-file-invoice-dollar"></i></div>
                <div class="kpi-value">${appState.quotes.length}</div>
                <div class="kpi-sub"><span>Clientes B2B activos: ${appState.clients.length}</span></div>
            </div>
            <div class="kpi-card">
                <div class="kpi-header"><span>Valor Total Cotizado</span><i class="fa-solid fa-sack-dollar"></i></div>
                <div class="kpi-value">${formatCRC(totalPortfolioRev)}</div>
                <div class="kpi-sub"><span>Utilidad proyectada: ${formatCRC(totalPortfolioProfit)}</span></div>
            </div>
            <div class="kpi-card">
                <div class="kpi-header"><span>Margen Promedio Cartera</span><i class="fa-solid fa-chart-pie"></i></div>
                <div class="kpi-value">${formatPct(avgPortfolioMargin, 1)}</div>
                <div class="kpi-sub"><span class="badge badge-success">Meta Corporativa Cumplida</span></div>
            </div>
            <div class="kpi-card kpi-info">
                <div class="kpi-header"><span>Cotización Activa (${escapeHtml(draft.id)})</span><i class="fa-solid fa-calculator"></i></div>
                <div class="kpi-value">${formatCRC(draftRevenue)}</div>
                <div class="kpi-sub"><span>Margen en vivo: <strong>${formatPct(draftMarginPct, 1)}</strong></span></div>
            </div>
        `;
    }

    const badgeQ = document.getElementById('badgeCountQuotes');
    if (badgeQ) badgeQ.textContent = String(appState.quotes.length);

    // Renderizar tarjetas de cotizaciones guardadas
    const listCont = document.getElementById('savedQuotesListContainer');
    if (listCont) {
        listCont.innerHTML = appState.quotes.map(q => {
            const rev = q.items.reduce((s, i) => s + i.cantidad * (i.precioFinal || i.precioLista), 0);
            const cost = q.items.reduce((s, i) => s + i.cantidad * i.costoUnit, 0);
            const marg = rev > 0 ? ((rev - cost) / rev) * 100 : 0;
            const isCurrent = q.id === draft.id;

            let statusBadge = `<span class="badge badge-info">${escapeHtml(q.estado)}</span>`;
            if (q.estado === 'APROBADA' || q.estado === 'CONVERTIDA_OC') {
                statusBadge = `<span class="badge badge-success">${escapeHtml(q.estado)}</span>`;
            }

            return `
                <div style="background:var(--bg-elevated); border:1px solid ${isCurrent ? 'var(--brand-green)' : 'var(--border-color)'}; border-radius:var(--radius-sm); padding:13px 15px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                        <div>
                            <strong style="color:var(--brand-green); font-size:13.5px;">${escapeHtml(q.id)}</strong>
                            <span style="font-size:11.5px; color:var(--text-muted); margin-left:6px;">${escapeHtml(q.fecha)}</span>
                            <div style="font-weight:700; font-size:14px; margin-top:2px;">${escapeHtml(q.clienteNombre)}</div>
                        </div>
                        ${statusBadge}
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center; font-size:12px; color:var(--text-secondary); margin:8px 0;">
                        <span>Lista: <strong>${escapeHtml(q.listaPrecio)}</strong> · ${q.items.length} SKUs</span>
                        <span>Monto: <strong style="color:var(--text-main);">${formatCRC(rev)}</strong> (Margen: <strong>${formatPct(marg, 1)}</strong>)</span>
                    </div>
                    <div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:8px;">
                        <button type="button" class="btn btn-secondary btn-sm btn-load-quote" data-qid="${escapeHtml(q.id)}">
                            <i class="fa-solid fa-pen-to-square"></i> Editar
                        </button>
                        <button type="button" class="btn btn-primary btn-sm btn-convert-quote-oc" data-qid="${escapeHtml(q.id)}">
                            <i class="fa-solid fa-truck-Arrow-right"></i> Convertir a Orden de Compra (OC)
                        </button>
                    </div>
                </div>
            `;
        }).join('');
    }
}

function loadQuoteIntoBuilder(quoteId, notify = true) {
    const found = appState.quotes.find(q => q.id === quoteId);
    if (!found) return;
    appState.activeQuoteDraft = structuredClone(found);
    renderQuotesModule();
    if (notify) showToast(`Cotización ${quoteId} cargada en el constructor`, 'fa-pen-to-square');
}

function openQuotePrintModal() {
    const draft = appState.activeQuoteDraft;
    const client = appState.clients.find(c => c.id === draft.clienteId) || appState.clients[0];
    const modal = document.getElementById('quotePrintModal');
    const content = document.getElementById('quotePrintableContent');
    if (!modal || !content) return;

    let totalRev = 0;
    const rowsHtml = draft.items.map(line => {
        const sub = line.cantidad * line.precioFinal;
        totalRev += sub;
        return `
            <tr style="border-bottom:1px solid #cbd5e1;">
                <td style="padding:8px;"><strong>${escapeHtml(line.sku)}</strong></td>
                <td style="padding:8px;">${escapeHtml(line.name)}</td>
                <td style="padding:8px; text-align:center;">${escapeHtml(line.unidad)}</td>
                <td style="padding:8px; text-align:right;">${formatNum(line.cantidad)}</td>
                <td style="padding:8px; text-align:right;">${formatCRC(line.precioLista)}</td>
                <td style="padding:8px; text-align:right;">${formatPct(line.descuentoPct, 1)}</td>
                <td style="padding:8px; text-align:right; font-weight:700;">${formatCRC(line.precioFinal)}</td>
                <td style="padding:8px; text-align:right; font-weight:700;">${formatCRC(sub)}</td>
            </tr>
        `;
    }).join('');

    content.innerHTML = `
        <div style="background:#ffffff; color:#0c2340; padding:26px; border-radius:10px; border:2px solid #2e8b57;">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:3px solid #2e8b57; padding-bottom:14px; margin-bottom:16px;">
                <div>
                    <h2 style="margin:0; color:#0c2340; font-size:22px; font-weight:800;">INTERFRUTD · EL FRUTO PERFECTO</h2>
                    <p style="margin:2px 0 0; font-size:12.5px; color:#475569;">Departamento Comercial & Pricing Corporativo · Santo Domingo, Heredia, Costa Rica</p>
                </div>
                <div style="text-align:right;">
                    <div style="font-size:18px; font-weight:800; color:#2e8b57;">${escapeHtml(draft.id)}</div>
                    <div style="font-size:12px; color:#475569;">Fecha: ${escapeHtml(draft.fecha)} · Vigencia: ${escapeHtml(draft.vigenciaHasta)}</div>
                </div>
            </div>

            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; background:#f8fafc; padding:12px 16px; border-radius:8px; margin-bottom:16px; font-size:13px;">
                <div>
                    <div><strong>Cliente:</strong> ${escapeHtml(client.nombre)}</div>
                    <div><strong>Contacto:</strong> ${escapeHtml(client.contacto)} (${escapeHtml(client.email)})</div>
                </div>
                <div>
                    <div><strong>Segmento / Lista:</strong> ${escapeHtml(client.segmento)} (${escapeHtml(draft.listaPrecio)})</div>
                    <div><strong>Condición Crédito:</strong> ${client.diasCredito} días · Ruta: ${escapeHtml(client.ruta)}</div>
                </div>
            </div>

            <table style="width:100%; border-collapse:collapse; font-size:12.5px; margin-bottom:16px;">
                <thead>
                    <tr style="background:#0c2340; color:#ffffff;">
                        <th style="padding:8px; text-align:left;">Código</th>
                        <th style="padding:8px; text-align:left;">Descripción del Producto</th>
                        <th style="padding:8px; text-align:center;">Und</th>
                        <th style="padding:8px; text-align:right;">Cant.</th>
                        <th style="padding:8px; text-align:right;">Precio Lista</th>
                        <th style="padding:8px; text-align:right;">Desc.</th>
                        <th style="padding:8px; text-align:right;">Precio Neto</th>
                        <th style="padding:8px; text-align:right;">Total Línea</th>
                    </tr>
                </thead>
                <tbody>${rowsHtml}</tbody>
            </table>

            <div style="display:flex; justify-content:space-between; align-items:center; border-top:2px solid #0c2340; padding-top:12px;">
                <div style="font-size:12px; color:#334155; max-width:60%;">
                    <strong>Observaciones Comerciales:</strong> ${escapeHtml(draft.notas || 'Entrega garantizada bajo cadena de frío.')}
                </div>
                <div style="text-align:right;">
                    <div style="font-size:12px; color:#475569;">Monto Total Ofertado</div>
                    <div style="font-size:22px; font-weight:800; color:#2e8b57;">${formatCRC(totalRev)}</div>
                </div>
            </div>
        </div>
    `;

    modal.classList.add('open');
}

// ============================================================================
// MÓDULO 3: ÓRDENES DE COMPRA (OC), VENTAS, INVENTARIOS Y ANÁLISIS ABC
// ============================================================================
function renderOrdersAndInventoryModule() {
    const baseCatalog = getFilteredCatalog();

    // Cálculos de Ventas, Inventarios y ABC
    const totalMonthlySales = baseCatalog.reduce((s, x) => s + x.ventasMensualesCrc, 0);
    const totalMonthlyUtil = baseCatalog.reduce((s, x) => s + x.utilidadMensualCrc, 0);
    const totalInvValue = baseCatalog.reduce((s, x) => s + (x.stockTotal * x.costoActual), 0);
    const lowCoverageCount = baseCatalog.filter(x => x.diasCobertura <= 5).length;

    // Cálculos de Órdenes de Compra (Fill Rate)
    let totalOrderedUnits = 0;
    let totalDispatchedUnits = 0;
    let totalOcAmount = 0;
    appState.purchaseOrders.forEach(oc => {
        oc.lineas.forEach(l => {
            totalOrderedUnits += l.cantPedida;
            totalDispatchedUnits += l.cantDespachada;
            totalOcAmount += l.cantDespachada * l.precioUnit;
        });
    });
    const globalFillRatePct = totalOrderedUnits > 0 ? (totalDispatchedUnits / totalOrderedUnits) * 100 : 100;

    const kpiGrid = document.getElementById('ordersKpiGrid');
    if (kpiGrid) {
        kpiGrid.innerHTML = `
            <div class="kpi-card">
                <div class="kpi-header"><span>Ventas Mensuales Proyectadas</span><i class="fa-solid fa-chart-line"></i></div>
                <div class="kpi-value">${formatCRC(totalMonthlySales)}</div>
                <div class="kpi-sub">
                    <span>Utilidad Bruta: <strong>${formatCRC(totalMonthlyUtil)}</strong></span>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-header"><span>Valorización de Inventario</span><i class="fa-solid fa-warehouse"></i></div>
                <div class="kpi-value">${formatCRC(totalInvValue)}</div>
                <div class="kpi-sub">
                    <span>CEDI 400 + Bodega 401 + Ruta 402</span>
                </div>
            </div>
            <div class="kpi-card ${globalFillRatePct < 95 ? 'kpi-warning' : ''}">
                <div class="kpi-header"><span>Nivel de Servicio (Fill Rate OC)</span><i class="fa-solid fa-truck-fast"></i></div>
                <div class="kpi-value">${formatPct(globalFillRatePct, 1)}</div>
                <div class="kpi-sub">
                    <span>${appState.purchaseOrders.length} OCs activas · Facturado: ${formatCRC(totalOcAmount)}</span>
                </div>
            </div>
            <div class="kpi-card ${lowCoverageCount > 0 ? 'kpi-danger' : ''}">
                <div class="kpi-header"><span>Alertas de Abastecimiento</span><i class="fa-solid fa-bell"></i></div>
                <div class="kpi-value">${lowCoverageCount} <span style="font-size:14px; color:var(--text-muted);">SKUs &le;5 días</span></div>
                <div class="kpi-sub">
                    <span>Requieren reposición inmediata a CEDI</span>
                </div>
            </div>
        `;
    }

    const badgeOC = document.getElementById('badgeCountOC');
    if (badgeOC) badgeOC.textContent = String(appState.purchaseOrders.length);

    // Tabla de Seguimiento de Órdenes de Compra (OC)
    const ocTbody = document.getElementById('purchaseOrdersTableBody');
    if (ocTbody) {
        const filteredOCs = appState.purchaseOrders.filter(oc =>
            appState.ocFilterStatus === 'ALL' ? true : oc.estado === appState.ocFilterStatus
        );

        ocTbody.innerHTML = filteredOCs.map(oc => {
            const pedida = oc.lineas.reduce((s, l) => s + l.cantPedida, 0);
            const desp = oc.lineas.reduce((s, l) => s + l.cantDespachada, 0);
            const monto = oc.lineas.reduce((s, l) => s + (l.cantDespachada * l.precioUnit), 0);
            const costo = oc.lineas.reduce((s, l) => s + (l.cantDespachada * l.costoUnit), 0);
            const margenOc = monto > 0 ? ((monto - costo) / monto) * 100 : 0;
            const fillRate = pedida > 0 ? (desp / pedida) * 100 : 100;

            let stateBadge = `<span class="badge badge-info">${escapeHtml(oc.estado)}</span>`;
            if (oc.estado === 'ENTREGADA') stateBadge = `<span class="badge badge-success"><i class="fa-solid fa-check-double"></i> ENTREGADA & FACTURADA</span>`;
            else if (oc.estado === 'EN_RUTA') stateBadge = `<span class="badge badge-info"><i class="fa-solid fa-truck-fast"></i> EN RUTA DESPACHO</span>`;
            else if (oc.estado === 'ALERTA_STOCK') stateBadge = `<span class="badge badge-danger"><i class="fa-solid fa-triangle-exclamation"></i> ALERTA STOCK PARCIAL</span>`;
            else if (oc.estado === 'EN_ALISTAMIENTO') stateBadge = `<span class="badge badge-warning"><i class="fa-solid fa-boxes-packing"></i> EN ALISTAMIENTO BODEGA</span>`;

            return `
                <tr>
                    <td><code style="font-weight:800; color:var(--brand-lime);">${escapeHtml(oc.ocNumero)}</code></td>
                    <td><span class="badge badge-neutral">${escapeHtml(oc.cotizacionRef)}</span></td>
                    <td>
                        <div style="font-weight:700;">${escapeHtml(oc.clienteNombre)}</div>
                        <div style="font-size:11px; color:var(--text-muted);">${escapeHtml(oc.segmento)} · ${oc.lineas.length} SKUs</div>
                    </td>
                    <td>${escapeHtml(oc.bodegaAsignada)}</td>
                    <td>${escapeHtml(oc.fechaEntregaCompromiso)}</td>
                    <td class="num-col">${formatNum(pedida)}</td>
                    <td class="num-col" style="font-weight:700;">${formatNum(desp)}</td>
                    <td class="num-col">
                        <span class="badge ${fillRate >= 98 ? 'badge-success' : (fillRate >= 90 ? 'badge-warning' : 'badge-danger')}">
                            ${formatPct(fillRate, 1)}
                        </span>
                    </td>
                    <td class="num-col" style="font-weight:700;">${formatCRC(monto)}</td>
                    <td class="num-col">${formatPct(margenOc, 1)}</td>
                    <td>${stateBadge}</td>
                    <td class="center-col">
                        <button type="button" class="btn btn-secondary btn-sm btn-advance-oc" data-oc="${escapeHtml(oc.ocNumero)}" title="Avanzar al siguiente estado del flujo comercial">
                            <i class="fa-solid fa-forward-step"></i> Avanzar
                        </button>
                    </td>
                </tr>
            `;
        }).join('');
    }

    // Tabla de Inventarios, Cobertura y Clasificación ABC
    const invTbody = document.getElementById('inventoryAbcTableBody');
    if (invTbody) {
        const invRows = baseCatalog.filter(item => {
            if (appState.inventoryFilter === 'REORDEN') return item.diasCobertura <= 5;
            if (appState.inventoryFilter === 'SOBRESTOCK') return item.diasCobertura > 25;
            return true;
        });

        invTbody.innerHTML = invRows.slice(0, 120).map(item => {
            let statusHtml = `<span class="badge badge-success">Cobertura Saludable</span>`;
            if (item.diasCobertura <= 5) {
                statusHtml = `<span class="badge badge-danger"><i class="fa-solid fa-bell"></i> Reorden Urgente (&le;5d)</span>`;
            } else if (item.diasCobertura > 25) {
                statusHtml = `<span class="badge badge-warning"><i class="fa-solid fa-box-archive"></i> Baja Rotación / Sobre-Stock</span>`;
            }

            return `
                <tr>
                    <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(item.sku)}</code></td>
                    <td style="font-weight:700;">${escapeHtml(item.name)}</td>
                    <td><span class="badge badge-neutral">${escapeHtml(item.categoria)}</span></td>
                    <td class="center-col"><span class="abc-badge abc-${item.abcClass}">${item.abcClass}</span></td>
                    <td class="num-col">${formatNum(item.stockCedi)}</td>
                    <td class="num-col">${formatNum(item.stockBodega401)}</td>
                    <td class="num-col">${formatNum(item.stockRuta402)}</td>
                    <td class="num-col" style="font-weight:800;">${formatNum(item.stockTotal)} ${escapeHtml(item.unidad)}</td>
                    <td class="num-col">${formatNum(item.demandaMensualUnd)}</td>
                    <td class="num-col" style="font-weight:700;">${formatCRC(item.ventasMensualesCrc)}</td>
                    <td class="num-col" style="font-weight:700; color:${item.diasCobertura <= 5 ? 'var(--danger)' : 'var(--text-main)'};">
                        ${formatNum(item.diasCobertura, 1)} días
                    </td>
                    <td class="num-col">${formatNum(item.puntoReorden)}</td>
                    <td>${statusHtml}</td>
                </tr>
            `;
        }).join('');
    }

    renderCommercialCharts(baseCatalog);
}

function renderCommercialCharts(catalogList) {
    if (typeof Chart === 'undefined') return;

    // Agrupar ventas y margen por Categoría
    const byCat = {};
    catalogList.forEach(item => {
        if (!byCat[item.categoria]) {
            byCat[item.categoria] = { sales: 0, profit: 0, invVal: 0 };
        }
        byCat[item.categoria].sales += item.ventasMensualesCrc;
        byCat[item.categoria].profit += item.utilidadMensualCrc;
        byCat[item.categoria].invVal += item.stockTotal * item.costoActual;
    });

    const catLabels = Object.keys(byCat).sort((a, b) => byCat[b].sales - byCat[a].sales);
    const salesData = catLabels.map(c => Math.round(byCat[c].sales / 1000)); // Miles ₡
    const marginData = catLabels.map(c => byCat[c].sales > 0 ? Math.round((byCat[c].profit / byCat[c].sales) * 1000) / 10 : 0);

    const ctx1 = document.getElementById('chartCategorySalesMargin');
    if (ctx1) {
        if (chartCatInstance) chartCatInstance.destroy();
        chartCatInstance = new Chart(ctx1, {
            type: 'bar',
            data: {
                labels: catLabels,
                datasets: [
                    {
                        label: 'Ventas Mensuales (Miles ₡)',
                        data: salesData,
                        backgroundColor: 'rgba(77, 184, 72, 0.72)',
                        borderColor: '#4db848',
                        borderWidth: 1,
                        borderRadius: 6,
                        yAxisID: 'y'
                    },
                    {
                        label: 'Margen Bruto (%)',
                        data: marginData,
                        type: 'line',
                        borderColor: '#a3e635',
                        backgroundColor: '#a3e635',
                        borderWidth: 2.5,
                        tension: 0.3,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { labels: { color: '#94a3b8', font: { family: 'Outfit', size: 11 } } }
                },
                scales: {
                    x: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { display: false } },
                    y: { position: 'left', ticks: { color: '#94a3b8' }, grid: { color: 'rgba(148,163,184,0.1)' } },
                    y1: { position: 'right', min: 0, max: 50, ticks: { color: '#a3e635', callback: v => v + '%' }, grid: { display: false } }
                }
            }
        });
    }

    // Gráfico 2: Clasificación ABC Ventas vs Valor de Inventario
    const abcGroups = { A: { sales: 0, inv: 0 }, B: { sales: 0, inv: 0 }, C: { sales: 0, inv: 0 } };
    catalogList.forEach(item => {
        const g = abcGroups[item.abcClass] || abcGroups.C;
        g.sales += item.ventasMensualesCrc;
        g.inv += item.stockTotal * item.costoActual;
    });

    const ctx2 = document.getElementById('chartAbcInventory');
    if (ctx2) {
        if (chartAbcInstance) chartAbcInstance.destroy();
        chartAbcInstance = new Chart(ctx2, {
            type: 'bar',
            data: {
                labels: ['Clase A (Top 80% Ingresos)', 'Clase B (15% Ingresos)', 'Clase C (5% Cola Larga)'],
                datasets: [
                    {
                        label: 'Facturación Mensual (Miles ₡)',
                        data: [
                            Math.round(abcGroups.A.sales / 1000),
                            Math.round(abcGroups.B.sales / 1000),
                            Math.round(abcGroups.C.sales / 1000)
                        ],
                        backgroundColor: 'rgba(34, 197, 94, 0.75)',
                        borderRadius: 6
                    },
                    {
                        label: 'Capital en Inventario (Miles ₡)',
                        data: [
                            Math.round(abcGroups.A.inv / 1000),
                            Math.round(abcGroups.B.inv / 1000),
                            Math.round(abcGroups.C.inv / 1000)
                        ],
                        backgroundColor: 'rgba(56, 189, 248, 0.75)',
                        borderRadius: 6
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { labels: { color: '#94a3b8', font: { family: 'Outfit', size: 11 } } }
                },
                scales: {
                    x: { ticks: { color: '#94a3b8' }, grid: { display: false } },
                    y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(148,163,184,0.1)' } }
                }
            }
        });
    }
}

// ============================================================================
// MÓDULO 4: QUERY BUILDER COMERCIAL VISUAL & EXCEL AVANZADO
// ============================================================================
function applyQueryPreset(presetName) {
    appState.qbPreset = presetName;
    if (presetName === 'criticos_abc_a') {
        appState.qbDataset = 'MASTER_JOIN';
        appState.qbFilterField = 'margenReposicionPct';
        appState.qbFilterOp = '<';
        appState.qbFilterVal = '22';
        appState.qbGroupBy = 'NONE';
        appState.qbOrderBy = 'ventasMensualesCrc_DESC';
    } else if (presetName === 'quiebre_stock') {
        appState.qbDataset = 'MASTER_JOIN';
        appState.qbFilterField = 'diasCobertura';
        appState.qbFilterOp = '<=';
        appState.qbFilterVal = '6';
        appState.qbGroupBy = 'NONE';
        appState.qbOrderBy = 'diasCobertura_ASC';
    } else if (presetName === 'brecha_supermercados') {
        appState.qbDataset = 'MASTER_JOIN';
        appState.qbFilterField = 'precioCompetenciaMin';
        appState.qbFilterOp = '>';
        appState.qbFilterVal = '0';
        appState.qbGroupBy = 'NONE';
        appState.qbOrderBy = 'ventasMensualesCrc_DESC';
    } else if (presetName === 'alza_costos') {
        appState.qbDataset = 'MASTER_JOIN';
        appState.qbFilterField = 'varCostoPct';
        appState.qbFilterOp = '>=';
        appState.qbFilterVal = '5';
        appState.qbGroupBy = 'NONE';
        appState.qbOrderBy = 'varCostoPct_DESC';
    } else if (presetName === 'group_by_categoria') {
        appState.qbDataset = 'MASTER_JOIN';
        appState.qbFilterField = 'ALL';
        appState.qbFilterOp = '>';
        appState.qbFilterVal = '';
        appState.qbGroupBy = 'categoria';
        appState.qbOrderBy = 'ventasMensualesCrc_DESC';
    } else if (presetName === 'catalogo_completo') {
        appState.qbDataset = 'MASTER_JOIN';
        appState.qbFilterField = 'ALL';
        appState.qbFilterOp = '>';
        appState.qbFilterVal = '';
        appState.qbGroupBy = 'NONE';
        appState.qbOrderBy = 'ventasMensualesCrc_DESC';
    }

    // Sincronizar selects del DOM
    const ds = document.getElementById('qbDatasetSelect');
    const ff = document.getElementById('qbFilterField');
    const fo = document.getElementById('qbFilterOperator');
    const fv = document.getElementById('qbFilterValue');
    const gb = document.getElementById('qbGroupBySelect');
    const ob = document.getElementById('qbOrderBySelect');

    if (ds) ds.value = appState.qbDataset;
    if (ff) ff.value = appState.qbFilterField;
    if (fo) fo.value = appState.qbFilterOp;
    if (fv) fv.value = appState.qbFilterVal;
    if (gb) gb.value = appState.qbGroupBy;
    if (ob) ob.value = appState.qbOrderBy;

    renderQueryBuilderModule();
}

function evaluateCondition(rowVal, op, targetRaw) {
    if (targetRaw === '' || targetRaw == null) return true;
    const numRow = Number(rowVal);
    const numTarget = Number(targetRaw);
    const bothNumeric = !isNaN(numRow) && !isNaN(numTarget) && String(targetRaw).trim() !== '';

    if (op === 'CONTAINS') {
        return String(rowVal ?? '').toLowerCase().includes(String(targetRaw).toLowerCase());
    }
    if (op === '=') {
        return bothNumeric ? numRow === numTarget : String(rowVal).toUpperCase() === String(targetRaw).trim().toUpperCase();
    }
    if (bothNumeric) {
        if (op === '<') return numRow < numTarget;
        if (op === '<=') return numRow <= numTarget;
        if (op === '>') return numRow > numTarget;
        if (op === '>=') return numRow >= numTarget;
    }
    return true;
}

function renderQueryBuilderModule() {
    const thead = document.getElementById('qbResultsThead');
    const tbody = document.getElementById('qbResultsTbody');
    const sqlBox = document.getElementById('qbSqlPreview');
    const statsBar = document.getElementById('qbExcelStatsBar');
    if (!thead || !tbody || !sqlBox || !statsBar) return;

    const ds = appState.qbDataset;
    const field = appState.qbFilterField;
    const op = appState.qbFilterOp;
    const val = appState.qbFilterVal;
    const groupBy = appState.qbGroupBy;
    const orderBy = appState.qbOrderBy;

    // Caso 1: Vista CENADA o Pipeline OCs
    if (ds === 'CENADA_VIEW') {
        sqlBox.textContent = `SELECT frumusaCod AS SKU, frumusaName AS Producto, frumusaUnit AS Presentacion, frumusaPrecioIva AS Precio_Interfrutd, cenadaModa AS Moda_CENADA, diffColones AS Brecha_CRC, verdict AS Posicion\nFROM vw_benchmark_mayorista_cenada\nORDER BY pctDiffModa ASC;`;
        thead.innerHTML = `
            <tr>
                <th>SKU</th><th>Producto Mayorista</th><th>Presentación</th>
                <th class="num-col">Precio Interfrutd</th><th class="num-col">Moda CENADA</th>
                <th class="num-col">Diferencia (₡)</th><th class="num-col">Var %</th><th>Veredicto</th>
            </tr>
        `;
        tbody.innerHTML = appState.cenadaList.map(r => `
            <tr>
                <td><code>${escapeHtml(r.frumusaCod)}</code></td>
                <td style="font-weight:700;">${escapeHtml(r.frumusaName)}</td>
                <td>${escapeHtml(r.frumusaUnit)}</td>
                <td class="num-col">${formatCRC(r.frumusaPrecioIva)}</td>
                <td class="num-col">${formatCRC(r.cenadaModa)}</td>
                <td class="num-col">${formatCRC(r.diffColones)}</td>
                <td class="num-col">${formatPct(r.pctDiffModa, 1)}</td>
                <td><span class="badge badge-info">${escapeHtml(r.verdict)}</span></td>
            </tr>
        `).join('');
        appState.lastQueryExportRows = appState.cenadaList;
        statsBar.innerHTML = `<span>Filas consultadas: <strong>${appState.cenadaList.length}</strong> presentaciones mayoristas vs PIMA CENADA</span>`;
        return;
    }

    if (ds === 'OC_PIPELINE') {
        sqlBox.textContent = `SELECT ocNumero, cotizacionRef, clienteNombre, bodegaAsignada, fechaEntregaCompromiso, estado, prioridad\nFROM tbl_ordenes_compra_b2b\nORDER BY fechaEntregaCompromiso ASC;`;
        thead.innerHTML = `
            <tr>
                <th>N° OC</th><th>Ref. Cotización</th><th>Cliente</th>
                <th>Bodega</th><th>Compromiso</th><th>Prioridad</th><th>Estado</th>
            </tr>
        `;
        tbody.innerHTML = appState.purchaseOrders.map(o => `
            <tr>
                <td><code>${escapeHtml(o.ocNumero)}</code></td>
                <td>${escapeHtml(o.cotizacionRef)}</td>
                <td style="font-weight:700;">${escapeHtml(o.clienteNombre)}</td>
                <td>${escapeHtml(o.bodegaAsignada)}</td>
                <td>${escapeHtml(o.fechaEntregaCompromiso)}</td>
                <td><span class="badge badge-warning">${escapeHtml(o.prioridad)}</span></td>
                <td><span class="badge badge-info">${escapeHtml(o.estado)}</span></td>
            </tr>
        `).join('');
        appState.lastQueryExportRows = appState.purchaseOrders;
        statsBar.innerHTML = `<span>Total Órdenes de Compra en Pipeline: <strong>${appState.purchaseOrders.length}</strong></span>`;
        return;
    }

    // Caso Principal: JOIN Maestro (Pricing + Ventas + Costos + Inventario + ABC)
    let rows = getFilteredCatalog().filter(item => {
        if (appState.qbPreset === 'criticos_abc_a' && field === 'margenReposicionPct' && val === '22') {
            return item.abcClass === 'A' && item.margenReposicionPct < 22;
        }
        if (field === 'ALL') return true;
        return evaluateCondition(item[field], op, val);
    });

    const [sortField, sortDir] = orderBy.split('_');
    rows.sort((a, b) => {
        const va = Number(a[sortField]) || 0;
        const vb = Number(b[sortField]) || 0;
        return sortDir === 'ASC' ? va - vb : vb - va;
    });

    const whereClause = field !== 'ALL' && val !== ''
        ? `\nWHERE ${field} ${op} '${val}'${appState.qbPreset === 'criticos_abc_a' ? " AND abcClass = 'A'" : ''}`
        : '';

    // Si hay GROUP BY activo
    if (groupBy !== 'NONE') {
        sqlBox.textContent = `SELECT ${groupBy}, COUNT(sku) AS Total_SKUs, SUM(stockTotal) AS Stock_Unidades, SUM(ventasMensualesCrc) AS Ventas_Mensuales_CRC, SUM(utilidadMensualCrc) AS Utilidad_CRC, ROUND(AVG(margenReposicionPct), 1) AS Margen_Promedio_Pct\nFROM vw_interfrutd_comercial_pricing${whereClause}\nGROUP BY ${groupBy}\nORDER BY Ventas_Mensuales_CRC DESC;`;

        const groups = {};
        rows.forEach(r => {
            const key = r[groupBy] || 'Sin Definir';
            if (!groups[key]) {
                groups[key] = { grupo: key, skus: 0, stock: 0, ventas: 0, utilidad: 0, sumMargin: 0 };
            }
            groups[key].skus += 1;
            groups[key].stock += r.stockTotal;
            groups[key].ventas += r.ventasMensualesCrc;
            groups[key].utilidad += r.utilidadMensualCrc;
            groups[key].sumMargin += r.margenReposicionPct;
        });

        const groupRows = Object.values(groups).map(g => ({
            Grupo: g.grupo,
            Total_SKUs: g.skus,
            Stock_Total: Math.round(g.stock),
            Ventas_Mensuales_CRC: g.ventas,
            Utilidad_Bruta_CRC: g.utilidad,
            Margen_Promedio_Pct: g.skus > 0 ? Math.round((g.sumMargin / g.skus) * 10) / 10 : 0
        })).sort((a, b) => b.Ventas_Mensuales_CRC - a.Ventas_Mensuales_CRC);

        appState.lastQueryExportRows = groupRows;

        thead.innerHTML = `
            <tr>
                <th>Grupo (${escapeHtml(groupBy)})</th>
                <th class="num-col">COUNT(SKUs)</th>
                <th class="num-col">SUM(Stock Total)</th>
                <th class="num-col">SUM(Ventas Mensuales)</th>
                <th class="num-col">SUM(Utilidad Bruta)</th>
                <th class="num-col">AVG(Margen Bruto %)</th>
            </tr>
        `;
        tbody.innerHTML = groupRows.map(g => `
            <tr>
                <td style="font-weight:800; color:var(--brand-green);">${escapeHtml(g.Grupo)}</td>
                <td class="num-col">${formatNum(g.Total_SKUs)}</td>
                <td class="num-col">${formatNum(g.Stock_Total)}</td>
                <td class="num-col" style="font-weight:700;">${formatCRC(g.Ventas_Mensuales_CRC)}</td>
                <td class="num-col" style="color:var(--success); font-weight:700;">${formatCRC(g.Utilidad_Bruta_CRC)}</td>
                <td class="num-col"><span class="badge badge-info">${formatPct(g.Margen_Promedio_Pct, 1)}</span></td>
            </tr>
        `).join('');

        const totalSalesG = groupRows.reduce((s, x) => s + x.Ventas_Mensuales_CRC, 0);
        statsBar.innerHTML = `
            <span>Grupos resultantes: <strong>${groupRows.length}</strong></span>
            <span>SKUs procesados: <strong>${rows.length}</strong></span>
            <span>Suma Ventas: <strong style="color:var(--brand-lime);">${formatCRC(totalSalesG)}</strong></span>
        `;
        return;
    }

    // Sin agrupar: Detalle de SKUs
    sqlBox.textContent = `SELECT sku, name, categoria, abcClass, costoReposicion, precioHoreca, precioDetalle, margenReposicionPct, precioCompetenciaMin, stockTotal, diasCobertura, ventasMensualesCrc\nFROM vw_interfrutd_comercial_pricing${whereClause}\nORDER BY ${sortField} ${sortDir};`;

    appState.lastQueryExportRows = rows.map(r => ({
        SKU: r.sku,
        Producto: r.name,
        Categoria: r.categoria,
        ABC: r.abcClass,
        Costo_Actual_CRC: r.costoActual,
        Costo_Reposicion_CRC: r.costoReposicion,
        Var_Costo_Pct: r.varCostoPct,
        Precio_Mayorista_CRC: r.precioMayorista,
        Precio_Horeca_CRC: r.precioHoreca,
        Precio_Detalle_CRC: r.precioDetalle,
        Margen_Real_Pct: r.margenReposicionPct,
        Min_Supermercados_CRC: r.precioCompetenciaMin,
        Stock_Total: r.stockTotal,
        Dias_Cobertura: r.diasCobertura,
        Ventas_Mensuales_CRC: r.ventasMensualesCrc,
        Utilidad_Mensual_CRC: r.utilidadMensualCrc
    }));

    thead.innerHTML = `
        <tr>
            <th>SKU</th>
            <th>Producto</th>
            <th>Categoría</th>
            <th class="center-col">ABC</th>
            <th class="num-col">Costo Repos.</th>
            <th class="num-col">Var Costo</th>
            <th class="num-col">Precio HORECA</th>
            <th class="num-col">Precio Detalle</th>
            <th class="num-col">Margen %</th>
            <th class="num-col">Mín. Superm.</th>
            <th class="num-col">Stock Total</th>
            <th class="num-col">Cobertura</th>
            <th class="num-col">Ventas Mensuales</th>
        </tr>
    `;

    tbody.innerHTML = rows.map(r => `
        <tr>
            <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(r.sku)}</code></td>
            <td style="font-weight:700;">${escapeHtml(r.name)}</td>
            <td><span class="badge badge-neutral">${escapeHtml(r.categoria)}</span></td>
            <td class="center-col"><span class="abc-badge abc-${r.abcClass}">${r.abcClass}</span></td>
            <td class="num-col">${formatCRC(r.costoReposicion)}</td>
            <td class="num-col">${r.varCostoPct > 0 ? '+' : ''}${formatPct(r.varCostoPct, 1)}</td>
            <td class="num-col">${formatCRC(r.precioHoreca)}</td>
            <td class="num-col" style="font-weight:700;">${formatCRC(r.precioDetalle)}</td>
            <td class="num-col">
                <span class="badge ${r.margenReposicionPct < 18 ? 'badge-danger' : 'badge-success'}">${formatPct(r.margenReposicionPct, 1)}</span>
            </td>
            <td class="num-col">${r.precioCompetenciaMin > 0 ? formatCRC(r.precioCompetenciaMin) : '—'}</td>
            <td class="num-col">${formatNum(r.stockTotal)}</td>
            <td class="num-col">${formatNum(r.diasCobertura, 1)} d</td>
            <td class="num-col" style="font-weight:700;">${formatCRC(r.ventasMensualesCrc)}</td>
        </tr>
    `).join('');

    const sumSales = rows.reduce((s, x) => s + x.ventasMensualesCrc, 0);
    const sumProfit = rows.reduce((s, x) => s + x.utilidadMensualCrc, 0);
    const avgMargin = rows.length > 0 ? rows.reduce((s, x) => s + x.margenReposicionPct, 0) / rows.length : 0;
    const sumStock = rows.reduce((s, x) => s + x.stockTotal, 0);

    statsBar.innerHTML = `
        <div style="display:flex; gap:16px; flex-wrap:wrap;">
            <span><i class="fa-solid fa-table"></i> Recuento: <strong>${rows.length} SKUs</strong></span>
            <span>Σ Ventas: <strong style="color:var(--brand-lime);">${formatCRC(sumSales)}</strong></span>
            <span>Σ Utilidad: <strong style="color:var(--success);">${formatCRC(sumProfit)}</strong></span>
            <span>x̄ Margen Promedio: <strong>${formatPct(avgMargin, 1)}</strong></span>
            <span>Σ Unidades en Stock: <strong>${formatNum(sumStock)}</strong></span>
        </div>
        <span style="color:var(--text-muted); font-size:11.5px;">Motor Query Builder SQL + SheetJS Activo</span>
    `;
}

// ============================================================================
// RENDERIZADO GLOBAL Y EVENTOS DE INTERFAZ
// ============================================================================
function renderActiveModule() {
    renderPricingModule();
    if (appState.activeTab === 'cotizaciones') renderQuotesModule();
    if (appState.activeTab === 'ordenes') renderOrdersAndInventoryModule();
    if (appState.activeTab === 'query-builder') renderQueryBuilderModule();
}

function setupEvents() {
    // Poblar selector global de categorías
    const catSelect = document.getElementById('globalCategoryFilter');
    if (catSelect) {
        const cats = [...new Set(appState.catalog.map(x => x.categoria))].sort();
        catSelect.innerHTML = `<option value="ALL">Todas las Categorías (${cats.length})</option>` +
            cats.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
        catSelect.addEventListener('change', () => {
            appState.globalCategory = catSelect.value;
            renderActiveModule();
        });
    }

    const abcSelect = document.getElementById('globalAbcFilter');
    if (abcSelect) {
        abcSelect.addEventListener('change', () => {
            appState.globalAbc = abcSelect.value;
            renderActiveModule();
        });
    }

    // Buscador global + Atajo Ctrl+K
    const searchInput = document.getElementById('globalSearchInput');
    if (searchInput) {
        searchInput.addEventListener('input', () => {
            appState.searchQuery = searchInput.value;
            renderActiveModule();
        });
        window.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
                e.preventDefault();
                searchInput.focus();
                searchInput.select();
            }
        });
    }

    // Navegación entre las 4 pestañas principales
    document.querySelectorAll('.module-tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.getAttribute('data-tab');
            appState.activeTab = tab;
            document.querySelectorAll('.module-tab-btn').forEach(b => {
                b.classList.toggle('active', b === btn);
                b.setAttribute('aria-selected', b === btn ? 'true' : 'false');
            });
            document.querySelectorAll('.tab-panel').forEach(p => {
                p.classList.toggle('active', p.id === `panel-${tab}`);
            });
            renderActiveModule();
        });
    });

    // Tema Claro / Oscuro Corporativo
    const btnTheme = document.getElementById('btnToggleTheme');
    if (btnTheme) {
        btnTheme.addEventListener('click', () => {
            const html = document.documentElement;
            const current = html.getAttribute('data-theme');
            const next = current === 'light' ? 'dark' : 'light';
            html.setAttribute('data-theme', next);
            const icon = document.getElementById('iconThemeToggle');
            if (icon) icon.className = next === 'light' ? 'fa-solid fa-moon' : 'fa-solid fa-sun';
            showToast(`Modo visual cambiado a ${next === 'light' ? 'Claro Corporativo' : 'Oscuro Ejecutivo'}`, 'fa-palette');
        });
    }

    // ------------------------------------------------------------------------
    // Eventos Módulo 1: Pricing & Listas
    // ------------------------------------------------------------------------
    const selChannel = document.getElementById('selectPriceListChannel');
    if (selChannel) {
        selChannel.addEventListener('change', () => {
            appState.activePriceListChannel = selChannel.value;
            renderPricingModule();
        });
    }

    const rangeMargin = document.getElementById('rangeTargetMargin');
    const labelMargin = document.getElementById('labelTargetMarginVal');
    if (rangeMargin) {
        rangeMargin.addEventListener('input', () => {
            appState.targetMarginSimPct = Number(rangeMargin.value);
            if (labelMargin) labelMargin.textContent = `${appState.targetMarginSimPct}%`;
            renderPricingModule();
        });
    }

    document.querySelectorAll('#pricingQuickChips .chip-btn').forEach(chip => {
        chip.addEventListener('click', () => {
            document.querySelectorAll('#pricingQuickChips .chip-btn').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            appState.pricingDiagFilter = chip.getAttribute('data-pfilter');
            renderPricingModule();
        });
    });

    const btnSwitchMatrix = document.getElementById('btnSwitchMatrixView');
    if (btnSwitchMatrix) {
        btnSwitchMatrix.addEventListener('click', () => {
            if (appState.pricingView === 'catalog') {
                appState.pricingView = 'cenada';
                btnSwitchMatrix.innerHTML = `<i class="fa-solid fa-table-list"></i> Volver a Matriz Maestra (320 SKUs)`;
            } else {
                appState.pricingView = 'catalog';
                btnSwitchMatrix.innerHTML = `<i class="fa-solid fa-warehouse"></i> Ver Comparativo Cajas PIMA CENADA (39)`;
            }
            renderPricingModule();
        });
    }

    // Edición en línea de precios y botón de enviar a cotización
    const pricingTbody = document.getElementById('pricingTableBody');
    if (pricingTbody) {
        pricingTbody.addEventListener('change', (e) => {
            if (e.target.classList.contains('inline-price-input')) {
                const sku = e.target.getAttribute('data-sku');
                const newPrice = Math.max(5, Number(e.target.value) || 0);
                const item = appState.catalog.find(x => x.sku === sku);
                if (item) {
                    item.precioDetalle = newPrice;
                    recalcSkuMetrics(item);
                    try {
                        const raw = localStorage.getItem('interfrutd_price_overrides');
                        const map = raw ? JSON.parse(raw) : {};
                        map[sku] = newPrice;
                        localStorage.setItem('interfrutd_price_overrides', JSON.stringify(map));
                    } catch (err) {}
                    renderPricingModule();
                    showToast(`Precio actualizado para SKU ${sku}: ${formatCRC(newPrice)} (Margen: ${formatPct(item.margenReposicionPct, 1)})`, 'fa-tag');
                }
            }
        });

        pricingTbody.addEventListener('click', (e) => {
            const btnQuote = e.target.closest('.btn-quick-quote');
            if (btnQuote) {
                const sku = btnQuote.getAttribute('data-sku');
                const item = appState.catalog.find(x => x.sku === sku);
                if (item) {
                    const listType = appState.activeQuoteDraft.listaPrecio || 'HORECA';
                    const precioLista = getBasePriceByList(item, listType);
                    const existing = appState.activeQuoteDraft.items.find(l => l.sku === sku);
                    if (existing) {
                        existing.cantidad += 25;
                    } else {
                        appState.activeQuoteDraft.items.push({
                            sku: item.sku,
                            name: item.name,
                            unidad: item.unidad,
                            cantidad: 30,
                            costoUnit: item.costoReposicion,
                            precioLista,
                            descuentoPct: 2,
                            precioFinal: Math.round(precioLista * 0.98)
                        });
                    }
                    showToast(`${item.name} agregado a la cotización ${appState.activeQuoteDraft.id}`, 'fa-cart-plus');
                    document.querySelector('.module-tab-btn[data-tab="cotizaciones"]')?.click();
                }
            }
        });
    }

    // Botón Optimizar Precios Críticos en 1 clic
    const btnOptimize = document.getElementById('btnApplySuggestedPrices');
    if (btnOptimize) {
        btnOptimize.addEventListener('click', () => {
            let count = 0;
            let monthlyGain = 0;
            appState.catalog.forEach(item => {
                if (item.margenReposicionPct < 16) {
                    const oldPrice = item.precioDetalle;
                    item.precioDetalle = item.precioSugeridoMeta;
                    monthlyGain += (item.precioDetalle - oldPrice) * item.demandaMensualUnd;
                    recalcSkuMetrics(item);
                    count++;
                }
            });
            renderPricingModule();
            if (count > 0) {
                showToast(`${count} SKUs críticos optimizados al margen meta (+${formatCRC(monthlyGain)}/mes de utilidad)`, 'fa-wand-magic-sparkles');
            } else {
                showToast('No hay SKUs con margen crítico (<16%) pendientes de optimizar.', 'fa-check');
            }
        });
    }

    const btnExportPricing = document.getElementById('btnExportPricingExcel');
    if (btnExportPricing) {
        btnExportPricing.addEventListener('click', () => {
            const rows = getFilteredCatalog().map(x => ({
                SKU: x.sku,
                Producto: x.name,
                Categoria: x.categoria,
                Unidad: x.unidad,
                ABC: x.abcClass,
                Costo_Actual_CRC: x.costoActual,
                Costo_Reposicion_CRC: x.costoReposicion,
                Var_Costo_Pct: x.varCostoPct,
                Lista_Mayorista_CRC: x.precioMayorista,
                Lista_Supermercado_CRC: x.precioSupermercado,
                Lista_HORECA_CRC: x.precioHoreca,
                Precio_Detalle_CRC: x.precioDetalle,
                Margen_Real_Pct: x.margenReposicionPct,
                Precio_Sugerido_Meta_CRC: x.precioSugeridoMeta,
                Min_Supermercados_CRC: x.precioCompetenciaMin,
                Estado_Pricing: x.estadoPrecio
            }));
            exportRowsToExcel(rows, 'Lista_Precios_Interfrutd', 'Interfrutd_Matriz_Precios_y_Costos.xlsx');
        });
    }

    // ------------------------------------------------------------------------
    // Eventos Módulo 2: Cotizador Comercial B2B
    // ------------------------------------------------------------------------
    const quoteClientSel = document.getElementById('quoteClientSelect');
    if (quoteClientSel) {
        quoteClientSel.addEventListener('change', () => {
            const cli = appState.clients.find(c => c.id === quoteClientSel.value);
            if (cli) {
                appState.activeQuoteDraft.clienteId = cli.id;
                appState.activeQuoteDraft.clienteNombre = cli.nombre;
                appState.activeQuoteDraft.segmento = cli.segmento;
                appState.activeQuoteDraft.listaPrecio = cli.listaAsignada;
                // Actualizar precios lista según la lista del cliente
                appState.activeQuoteDraft.items.forEach(line => {
                    const catSku = appState.catalog.find(x => x.sku === line.sku);
                    if (catSku) {
                        line.precioLista = getBasePriceByList(catSku, cli.listaAsignada);
                        line.descuentoPct = cli.descuentoBasePct;
                    }
                });
                renderQuotesModule();
            }
        });
    }

    const quoteListSel = document.getElementById('quotePriceListSelect');
    if (quoteListSel) {
        quoteListSel.addEventListener('change', () => {
            const newList = quoteListSel.value;
            appState.activeQuoteDraft.listaPrecio = newList;
            appState.activeQuoteDraft.items.forEach(line => {
                const catSku = appState.catalog.find(x => x.sku === line.sku);
                if (catSku) line.precioLista = getBasePriceByList(catSku, newList);
            });
            renderQuotesModule();
        });
    }

    const btnApplyGlobalDisc = document.getElementById('btnApplyGlobalDiscount');
    if (btnApplyGlobalDisc) {
        btnApplyGlobalDisc.addEventListener('click', () => {
            const d = Math.min(35, Math.max(0, Number(document.getElementById('quoteGlobalDiscount')?.value) || 0));
            appState.activeQuoteDraft.items.forEach(l => { l.descuentoPct = d; });
            renderQuotesModule();
            showToast(`Descuento global de ${d}% aplicado a todas las líneas`, 'fa-percent');
        });
    }

    const btnAddLine = document.getElementById('btnAddQuoteLine');
    if (btnAddLine) {
        btnAddLine.addEventListener('click', () => {
            const sku = document.getElementById('quoteAddSkuSelect')?.value;
            const qty = Math.max(1, Number(document.getElementById('quoteAddQtyInput')?.value) || 25);
            const item = appState.catalog.find(x => x.sku === sku);
            if (!item) return;

            const listType = appState.activeQuoteDraft.listaPrecio || 'HORECA';
            const precioLista = getBasePriceByList(item, listType);
            const existing = appState.activeQuoteDraft.items.find(l => l.sku === sku);
            if (existing) {
                existing.cantidad += qty;
            } else {
                appState.activeQuoteDraft.items.push({
                    sku: item.sku,
                    name: item.name,
                    unidad: item.unidad,
                    cantidad: qty,
                    costoUnit: item.costoReposicion,
                    precioLista,
                    descuentoPct: Number(document.getElementById('quoteGlobalDiscount')?.value) || 0,
                    precioFinal: precioLista
                });
            }
            renderQuotesModule();
            showToast(`${item.name} (${qty} ${item.unidad}) agregado a la cotización`, 'fa-cart-plus');
        });
    }

    const quoteLinesTbody = document.getElementById('quoteLinesTableBody');
    if (quoteLinesTbody) {
        quoteLinesTbody.addEventListener('change', (e) => {
            const idx = Number(e.target.getAttribute('data-idx'));
            const line = appState.activeQuoteDraft.items[idx];
            if (!line) return;
            if (e.target.classList.contains('quote-line-qty')) {
                line.cantidad = Math.max(1, Number(e.target.value) || 1);
            } else if (e.target.classList.contains('quote-line-disc')) {
                line.descuentoPct = Math.min(40, Math.max(0, Number(e.target.value) || 0));
            }
            renderQuotesModule();
        });

        quoteLinesTbody.addEventListener('click', (e) => {
            const btnRem = e.target.closest('.btn-remove-qline');
            if (btnRem) {
                const idx = Number(btnRem.getAttribute('data-idx'));
                appState.activeQuoteDraft.items.splice(idx, 1);
                renderQuotesModule();
            }
        });
    }

    const btnNewQuote = document.getElementById('btnNewBlankQuote');
    if (btnNewQuote) {
        btnNewQuote.addEventListener('click', () => {
            const nextNum = 142 + appState.quotes.length + 1;
            const firstClient = appState.clients[0];
            appState.activeQuoteDraft = {
                id: `COT-2026-0${nextNum}`,
                fecha: new Date().toISOString().slice(0, 10),
                vigenciaHasta: '2026-11-10',
                clienteId: firstClient.id,
                clienteNombre: firstClient.nombre,
                segmento: firstClient.segmento,
                listaPrecio: firstClient.listaAsignada,
                estado: 'BORRADOR',
                notas: 'Oferta comercial sujeta a confirmación de orden de compra.',
                items: []
            };
            renderQuotesModule();
            showToast(`Nueva cotización ${appState.activeQuoteDraft.id} iniciada`, 'fa-file-circle-plus');
        });
    }

    const btnSaveQuote = document.getElementById('btnSaveQuote');
    if (btnSaveQuote) {
        btnSaveQuote.addEventListener('click', () => {
            const draft = appState.activeQuoteDraft;
            if (draft.items.length === 0) {
                showToast('Agregue al menos 1 producto antes de guardar la cotización.', 'fa-triangle-exclamation');
                return;
            }
            const cli = appState.clients.find(c => c.id === draft.clienteId) || appState.clients[0];
            draft.clienteNombre = cli.nombre;
            draft.segmento = cli.segmento;
            draft.vigenciaHasta = document.getElementById('quoteValidUntil')?.value || draft.vigenciaHasta;
            draft.notas = document.getElementById('quoteNotesInput')?.value || '';
            draft.estado = 'ENVIADA';

            const idx = appState.quotes.findIndex(q => q.id === draft.id);
            if (idx >= 0) {
                appState.quotes[idx] = structuredClone(draft);
            } else {
                appState.quotes.unshift(structuredClone(draft));
            }
            try {
                localStorage.setItem('interfrutd_quotes', JSON.stringify(appState.quotes));
            } catch (e) {}
            renderQuotesModule();
            showToast(`Cotización ${draft.id} guardada exitosamente en cartera`, 'fa-floppy-disk');
        });
    }

    const btnPreviewPdf = document.getElementById('btnPreviewQuotePdf');
    if (btnPreviewPdf) btnPreviewPdf.addEventListener('click', openQuotePrintModal);

    const btnCloseModal = document.getElementById('btnCloseQuoteModal');
    if (btnCloseModal) {
        btnCloseModal.addEventListener('click', () => {
            document.getElementById('quotePrintModal')?.classList.remove('open');
        });
    }

    const btnPrintNow = document.getElementById('btnPrintQuoteNow');
    if (btnPrintNow) {
        btnPrintNow.addEventListener('click', () => window.print());
    }

    const savedQuotesCont = document.getElementById('savedQuotesListContainer');
    if (savedQuotesCont) {
        savedQuotesCont.addEventListener('click', (e) => {
            const btnLoad = e.target.closest('.btn-load-quote');
            if (btnLoad) {
                loadQuoteIntoBuilder(btnLoad.getAttribute('data-qid'));
                return;
            }
            const btnConvert = e.target.closest('.btn-convert-quote-oc');
            if (btnConvert) {
                const qid = btnConvert.getAttribute('data-qid');
                const quote = appState.quotes.find(q => q.id === qid);
                if (!quote) return;

                quote.estado = 'CONVERTIDA_OC';
                const newOcNumber = `OC-${88440 + appState.purchaseOrders.length}`;
                const newOC = {
                    ocNumero: newOcNumber,
                    cotizacionRef: quote.id,
                    fechaRecepcion: new Date().toISOString().slice(0, 10),
                    fechaEntregaCompromiso: quote.vigenciaHasta,
                    clienteNombre: quote.clienteNombre,
                    segmento: quote.segmento || 'Corporativo',
                    canal: quote.listaPrecio,
                    estado: 'OC_RECIBIDA',
                    prioridad: 'ALTA',
                    bodegaAsignada: 'Bodega Comercial 401',
                    lineas: quote.items.map(it => ({
                        sku: it.sku,
                        name: it.name,
                        cantPedida: it.cantidad,
                        cantDespachada: it.cantidad,
                        precioUnit: it.precioFinal || it.precioLista,
                        costoUnit: it.costoUnit
                    }))
                };
                appState.purchaseOrders.unshift(newOC);
                try {
                    localStorage.setItem('interfrutd_purchase_orders', JSON.stringify(appState.purchaseOrders));
                    localStorage.setItem('interfrutd_quotes', JSON.stringify(appState.quotes));
                } catch (err) {}

                renderQuotesModule();
                showToast(`Cotización ${quote.id} convertida en Orden de Compra ${newOcNumber}`, 'fa-truck-fast');
                document.querySelector('.module-tab-btn[data-tab="ordenes"]')?.click();
            }
        });
    }

    const btnExportQuotes = document.getElementById('btnExportQuotesExcel');
    if (btnExportQuotes) {
        btnExportQuotes.addEventListener('click', () => {
            const flat = [];
            appState.quotes.forEach(q => {
                q.items.forEach(it => {
                    flat.push({
                        Cotizacion_ID: q.id,
                        Fecha: q.fecha,
                        Cliente: q.clienteNombre,
                        Lista: q.listaPrecio,
                        Estado: q.estado,
                        SKU: it.sku,
                        Producto: it.name,
                        Cantidad: it.cantidad,
                        Costo_Unit_CRC: it.costoUnit,
                        Precio_Ofertado_CRC: it.precioFinal,
                        Subtotal_CRC: it.cantidad * it.precioFinal
                    });
                });
            });
            exportRowsToExcel(flat, 'Cotizaciones_B2B', 'Interfrutd_Cartera_Cotizaciones_B2B.xlsx');
        });
    }

    // ------------------------------------------------------------------------
    // Eventos Módulo 3: Órdenes de Compra & Inventarios
    // ------------------------------------------------------------------------
    document.querySelectorAll('#ocStatusFilterGroup .chip-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#ocStatusFilterGroup .chip-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            appState.ocFilterStatus = btn.getAttribute('data-ocfilter');
            renderOrdersAndInventoryModule();
        });
    });

    const ocTbody = document.getElementById('purchaseOrdersTableBody');
    if (ocTbody) {
        ocTbody.addEventListener('click', (e) => {
            const btnAdv = e.target.closest('.btn-advance-oc');
            if (btnAdv) {
                const ocNum = btnAdv.getAttribute('data-oc');
                const oc = appState.purchaseOrders.find(x => x.ocNumero === ocNum);
                if (!oc) return;
                const flow = ['OC_RECIBIDA', 'EN_ALISTAMIENTO', 'EN_RUTA', 'ENTREGADA'];
                const currIdx = flow.indexOf(oc.estado);
                oc.estado = currIdx >= 0 && currIdx < flow.length - 1 ? flow[currIdx + 1] : 'ENTREGADA';
                try {
                    localStorage.setItem('interfrutd_purchase_orders', JSON.stringify(appState.purchaseOrders));
                } catch (err) {}
                renderOrdersAndInventoryModule();
                showToast(`Orden ${oc.ocNumero} actualizada a estado: ${oc.estado}`, 'fa-truck-fast');
            }
        });
    }

    document.querySelectorAll('[data-invfilter]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('[data-invfilter]').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            appState.inventoryFilter = btn.getAttribute('data-invfilter');
            renderOrdersAndInventoryModule();
        });
    });

    const btnExportInv = document.getElementById('btnExportInventoryExcel');
    if (btnExportInv) {
        btnExportInv.addEventListener('click', () => {
            const rows = getFilteredCatalog().map(x => ({
                SKU: x.sku,
                Producto: x.name,
                Categoria: x.categoria,
                Clasificacion_ABC: x.abcClass,
                Stock_CEDI_400: x.stockCedi,
                Stock_Bodega_401: x.stockBodega401,
                Stock_Ruta_402: x.stockRuta402,
                Stock_Total: x.stockTotal,
                Demanda_Mensual_Und: x.demandaMensualUnd,
                Ventas_Mensuales_CRC: x.ventasMensualesCrc,
                Dias_Cobertura: x.diasCobertura,
                Punto_Reorden: x.puntoReorden
            }));
            exportRowsToExcel(rows, 'Inventario_y_ABC', 'Interfrutd_Inventario_Cobertura_ABC.xlsx');
        });
    }

    // ------------------------------------------------------------------------
    // Eventos Módulo 4: Query Builder Comercial
    // ------------------------------------------------------------------------
    document.querySelectorAll('#qbPresetButtons .chip-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#qbPresetButtons .chip-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            applyQueryPreset(btn.getAttribute('data-preset'));
        });
    });

    ['qbDatasetSelect', 'qbFilterField', 'qbFilterOperator', 'qbFilterValue', 'qbGroupBySelect', 'qbOrderBySelect'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', () => {
                appState.qbDataset = document.getElementById('qbDatasetSelect')?.value || 'MASTER_JOIN';
                appState.qbFilterField = document.getElementById('qbFilterField')?.value || 'ALL';
                appState.qbFilterOp = document.getElementById('qbFilterOperator')?.value || '<';
                appState.qbFilterVal = document.getElementById('qbFilterValue')?.value || '';
                appState.qbGroupBy = document.getElementById('qbGroupBySelect')?.value || 'NONE';
                appState.qbOrderBy = document.getElementById('qbOrderBySelect')?.value || 'ventasMensualesCrc_DESC';
                renderQueryBuilderModule();
            });
        }
    });

    const btnCopySql = document.getElementById('btnCopyGeneratedSql');
    if (btnCopySql) {
        btnCopySql.addEventListener('click', () => {
            const sql = document.getElementById('qbSqlPreview')?.textContent || '';
            navigator.clipboard?.writeText(sql);
            showToast('Consulta SQL copiada al portapapeles', 'fa-copy');
        });
    }

    const btnExportQb = document.getElementById('btnExportQueryExcel');
    if (btnExportQb) {
        btnExportQb.addEventListener('click', () => {
            exportRowsToExcel(appState.lastQueryExportRows, 'Query_Builder_Result', 'Interfrutd_QueryBuilder_Export.xlsx');
        });
    }

    // ------------------------------------------------------------------------
    // Importación y Exportación Corporativa Completa (.xlsx)
    // ------------------------------------------------------------------------
    const btnUploadExcel = document.getElementById('btnUploadExcelCatalog');
    const inputExcel = document.getElementById('inputExcelUpload');
    if (btnUploadExcel && inputExcel) {
        btnUploadExcel.addEventListener('click', () => inputExcel.click());
        inputExcel.addEventListener('change', async (e) => {
            const file = e.target.files?.[0];
            if (!file || typeof XLSX === 'undefined') return;
            try {
                const buf = await file.arrayBuffer();
                const wb = XLSX.read(buf, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws);
                let updated = 0;
                rows.forEach(r => {
                    const code = String(r.SKU || r.sku || r.CODIGO || r.Codigo || r.ARTICULO || '').trim().toUpperCase();
                    if (!code) return;
                    const found = appState.catalog.find(x => x.sku === code);
                    if (found) {
                        const newCost = Number(r.Costo_Reposicion_CRC || r.COSTO || r.Costo || 0);
                        const newPrice = Number(r.Precio_Detalle_CRC || r.PRECIO || r.Precio || 0);
                        if (newCost > 0) found.costoReposicion = newCost;
                        if (newPrice > 0) found.precioDetalle = newPrice;
                        recalcSkuMetrics(found);
                        updated++;
                    }
                });
                renderActiveModule();
                showToast(`Archivo "${file.name}" procesado: ${updated} SKUs actualizados en vivo.`, 'fa-file-excel');
            } catch (err) {
                showToast('Error al leer el archivo Excel.', 'fa-circle-xmark');
            }
        });
    }

    const btnExportWorkbook = document.getElementById('btnExportFullWorkbook');
    if (btnExportWorkbook) {
        btnExportWorkbook.addEventListener('click', () => {
            exportMultiSheetExcel([
                {
                    name: '1_Pricing_y_Listas',
                    data: appState.catalog.map(x => ({
                        SKU: x.sku,
                        Producto: x.name,
                        Categoria: x.categoria,
                        ABC: x.abcClass,
                        Costo_Actual: x.costoActual,
                        Costo_Reposicion: x.costoReposicion,
                        Lista_Mayorista: x.precioMayorista,
                        Lista_Supermercado: x.precioSupermercado,
                        Lista_HORECA: x.precioHoreca,
                        Precio_Detalle: x.precioDetalle,
                        Margen_Pct: x.margenReposicionPct,
                        Min_Supermercados: x.precioCompetenciaMin
                    }))
                },
                {
                    name: '2_Benchmark_CENADA',
                    data: appState.cenadaList
                },
                {
                    name: '3_Ordenes_Compra_OC',
                    data: appState.purchaseOrders.map(o => ({
                        OC: o.ocNumero,
                        Cotizacion_Ref: o.cotizacionRef,
                        Cliente: o.clienteNombre,
                        Bodega: o.bodegaAsignada,
                        Entrega: o.fechaEntregaCompromiso,
                        Estado: o.estado
                    }))
                }
            ], 'Reporte_Corporativo_Interfrutd_Pricing_Comercial.xlsx');
        });
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initStateData();
    setupEvents();
    applyQueryPreset('criticos_abc_a');
    renderActiveModule();
});
