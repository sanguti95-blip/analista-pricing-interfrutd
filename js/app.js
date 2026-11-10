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
    channelMargins: {
        SUPERMERCADO: 24,
        HORECA: 28,
        MAYORISTA: 18,
        DETALLE: 32
    },
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

    // Módulo 3: Programación de Días, Clientes y Órdenes de Compra (OC)
    selectedScheduleDay: 'ALL',
    purchaseOrders: [],
    ocFilterStatus: 'ALL',
    inventoryFilter: 'ALL'
};

let chartCatInstance = null;
let chartAbcInstance = null;

// ============================================================================
// INICIALIZACIÓN Y RESTAURACIÓN DE ESTADO LOCAL
// ============================================================================
function syncChannelMarginInputs() {
    const map = [
        { key: 'SUPERMERCADO', rId: 'rangeMarginSupermercado', nId: 'numMarginSupermercado' },
        { key: 'HORECA', rId: 'rangeMarginHoreca', nId: 'numMarginHoreca' },
        { key: 'MAYORISTA', rId: 'rangeMarginMayorista', nId: 'numMarginMayorista' },
        { key: 'DETALLE', rId: 'rangeMarginDetalle', nId: 'numMarginDetalle' }
    ];
    map.forEach(({ key, rId, nId }) => {
        const val = appState.channelMargins[key];
        const rEl = document.getElementById(rId);
        const nEl = document.getElementById(nId);
        if (rEl) rEl.value = val;
        if (nEl) nEl.value = val;
    });
}

function initStateData() {
    // Restaurar márgenes personalizados por tipo de cadena
    try {
        const savedMargins = localStorage.getItem('interfrutd_channel_margins');
        if (savedMargins) {
            appState.channelMargins = { ...appState.channelMargins, ...JSON.parse(savedMargins) };
        }
    } catch (e) {}
    syncChannelMarginInputs();

    // Clonar catálogo base y aplicar precios/costos personalizados si existen en localStorage
    let savedPriceOverrides = {};
    let savedCostOverrides = {};
    try {
        const raw = localStorage.getItem('interfrutd_price_overrides');
        if (raw) savedPriceOverrides = JSON.parse(raw) || {};
        const rawCost = localStorage.getItem('interfrutd_cost_overrides');
        if (rawCost) savedCostOverrides = JSON.parse(rawCost) || {};
    } catch (e) {}

    appState.catalog = INTERFRUTD_CATALOG.map(item => {
        const copy = { ...item };
        if (savedCostOverrides[copy.sku] !== undefined && Number(savedCostOverrides[copy.sku]) > 0) {
            copy.costoReposicion = Number(savedCostOverrides[copy.sku]);
        }
        if (savedPriceOverrides[copy.sku] !== undefined && Number(savedPriceOverrides[copy.sku]) > 0 && Number(savedPriceOverrides[copy.sku]) < 500000) {
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

    // Restaurar clientes B2B (creados o editados) y su programación personalizada de días
    try {
        const savedClients = localStorage.getItem('interfrutd_clients_v3');
        if (savedClients) {
            appState.clients = JSON.parse(savedClients);
        } else {
            const savedClientDays = localStorage.getItem('interfrutd_client_days');
            const daysMap = savedClientDays ? JSON.parse(savedClientDays) : {};
            appState.clients = B2B_CLIENTS.map(c => ({
                ...c,
                diasEntrega: Array.isArray(daysMap[c.id]) ? daysMap[c.id] : [...(c.diasEntrega || ['L', 'K', 'V'])]
            }));
        }
    } catch (e) {
        appState.clients = structuredClone(B2B_CLIENTS);
    }

    // Órdenes de compra
    try {
        const savedOC = localStorage.getItem('interfrutd_purchase_orders_v2');
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
    item.varCostoPct = item.costoActual > 0 ? Math.round(((cost - item.costoActual) / item.costoActual) * 1000) / 10 : 0;

    const mSuper = (appState.channelMargins.SUPERMERCADO || 24) / 100;
    const mHoreca = (appState.channelMargins.HORECA || 28) / 100;
    const mMayorista = (appState.channelMargins.MAYORISTA || 18) / 100;
    const mDetalle = (appState.channelMargins.DETALLE || 32) / 100;

    item.precioMayorista = Math.max(10, Math.round((cost / (1 - mMayorista)) / 5) * 5);
    item.precioSupermercado = Math.max(15, Math.round((cost / (1 - mSuper)) / 5) * 5);
    item.precioHoreca = Math.max(20, Math.round((cost / (1 - mHoreca)) / 5) * 5);
    item.precioSugeridoMeta = Math.max(15, Math.round((cost / (1 - mDetalle)) / 5) * 5);

    if (item.precioDetalle >= 500000 || item.precioDetalle <= 0) {
        item.precioDetalle = item.precioSugeridoMeta;
    }

    const price = item.precioDetalle;
    item.margenReposicionPct = Math.round(((price - cost) / price) * 1000) / 10;
    item.margenActualPct = Math.round(((price - item.costoActual) / price) * 1000) / 10;
    item.ventasMensualesCrc = Math.round(item.demandaMensualUnd * price);
    item.utilidadMensualCrc = Math.round(item.demandaMensualUnd * (price - cost));

    if (item.margenReposicionPct < 16) item.estadoPrecio = 'MARGEN_CRITICO';
    else if (item.varCostoPct >= 6) item.estadoPrecio = 'COSTO_SUBIO';
    else if (item.precioCompetenciaMin > 0 && price > item.precioCompetenciaMin * 1.15) item.estadoPrecio = 'CARO_VS_COMPETENCIA';
    else if (item.precioCompetenciaMin > 0 && price < item.precioCompetenciaMin * 0.85 && item.margenReposicionPct < appState.channelMargins.DETALLE) item.estadoPrecio = 'OPORTUNIDAD_PRECIO';
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
                    <span>Meta Detalle: <strong>${appState.channelMargins.DETALLE}%</strong></span>
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
    const quickChips = document.getElementById('pricingQuickChips');
    const btnOptimize = document.getElementById('btnApplySuggestedPrices');
    if (appState.pricingView === 'cenada') {
        if (catalogCont) catalogCont.style.display = 'none';
        if (quickChips) quickChips.style.visibility = 'hidden';
        if (btnOptimize) btnOptimize.style.display = 'none';
        if (cenadaCont) cenadaCont.style.display = 'block';
        renderCenadaComparisonTable();
        return;
    } else {
        if (catalogCont) catalogCont.style.display = 'block';
        if (quickChips) quickChips.style.visibility = 'visible';
        if (btnOptimize) btnOptimize.style.display = '';
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

    tbody.innerHTML = rows.map(item => {
        const costVal = Math.round(item.costoReposicion > 0 ? item.costoReposicion : item.costoActual);
        const varInline = item.varCostoPct >= 6
            ? `<span class="badge badge-danger" style="margin-left:4px; font-size:10px;">+${formatPct(item.varCostoPct, 0)}</span>`
            : (item.varCostoPct < 0
                ? `<span class="badge badge-success" style="margin-left:4px; font-size:10px;">${formatPct(item.varCostoPct, 0)}</span>`
                : '');

        let marginBadgeClass = 'badge-success';
        if (item.margenReposicionPct < 16) marginBadgeClass = 'badge-danger';
        else if (item.margenReposicionPct < 24) marginBadgeClass = 'badge-warning';

        let diagHtml = `<span class="badge badge-success"><i class="fa-solid fa-check"></i> Óptimo</span>`;
        if (item.estadoPrecio === 'MARGEN_CRITICO') {
            diagHtml = `<span class="badge badge-danger"><i class="fa-solid fa-bolt"></i> Crítico</span>`;
        } else if (item.estadoPrecio === 'COSTO_SUBIO') {
            diagHtml = `<span class="badge badge-warning"><i class="fa-solid fa-arrow-trend-up"></i> Alza Costo</span>`;
        } else if (item.estadoPrecio === 'OPORTUNIDAD_PRECIO') {
            diagHtml = `<span class="badge badge-info"><i class="fa-solid fa-circle-up"></i> Oportunidad</span>`;
        } else if (item.estadoPrecio === 'CARO_VS_COMPETENCIA') {
            diagHtml = `<span class="badge badge-warning">Sobre Mkt</span>`;
        }

        const superHint = item.precioCompetenciaMin > 0
            ? `<div style="font-size:11px; color:var(--text-muted);" title="Ref. Mín. Supermercados: ${formatCRC(item.precioCompetenciaMin)}">Ref: ${formatCRC(item.precioCompetenciaMin)}</div>`
            : '';

        return `
            <tr>
                <td style="white-space:nowrap;">
                    <code style="font-weight:700; color:var(--brand-green); margin-right:6px;">${escapeHtml(item.sku)}</code>
                    <span class="abc-badge abc-${item.abcClass}">${item.abcClass}</span>
                </td>
                <td>
                    <div style="font-weight:700;">${escapeHtml(item.name)}</div>
                    <div style="font-size:11px; color:var(--text-muted);">${escapeHtml(item.categoria)} · ${escapeHtml(item.unidad)} (x${item.factorEmpaque})</div>
                </td>
                <td class="num-col" style="white-space:nowrap;">
                    <div style="display:inline-flex; align-items:center;">
                        <input type="number" class="inline-price-input inline-cost-input" data-sku="${escapeHtml(item.sku)}" value="${costVal}" min="5" step="5" style="width:78px;" title="Editar Costo de Reposición (COGS)" aria-label="Costo reposición COGS para SKU ${escapeHtml(item.sku)}">
                        ${varInline}
                    </div>
                </td>
                <td class="num-col">
                    <div style="font-weight:700;">${formatCRC(item.precioSupermercado)}</div>
                    ${superHint}
                </td>
                <td class="num-col" style="font-weight:700;">${formatCRC(item.precioHoreca)}</td>
                <td class="num-col" style="color:var(--text-secondary);">${formatCRC(item.precioMayorista)}</td>
                <td class="num-col" style="white-space:nowrap;">
                    <div style="display:inline-flex; align-items:center; gap:6px;">
                        <input type="number" class="inline-price-input inline-retail-input" data-sku="${escapeHtml(item.sku)}" value="${item.precioDetalle}" min="5" step="5" aria-label="Precio detalle para SKU ${escapeHtml(item.sku)}">
                        <span class="badge ${marginBadgeClass}">${formatPct(item.margenReposicionPct, 1)}</span>
                    </div>
                </td>
                <td class="center-col" style="white-space:nowrap;">
                    <div style="display:inline-flex; align-items:center; gap:6px;">
                        ${diagHtml}
                        <button type="button" class="btn btn-secondary btn-sm btn-quick-quote" data-sku="${escapeHtml(item.sku)}" title="Agregar a Cotización B2B">
                            <i class="fa-solid fa-plus"></i>
                        </button>
                    </div>
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
        const catMatch = appState.catalog.find(x => x.sku === String(r.frumusaCod));
        const liveInterfrutdPrice = catMatch ? catMatch.precioMayorista : Number(r.frumusaPrecioIva) || 0;
        const cenadaPrice = Number(r.cenadaModa) || 0;
        const diff = liveInterfrutdPrice - cenadaPrice;
        const pct = cenadaPrice > 0 ? (diff / cenadaPrice) * 100 : 0;

        let badge = `<span class="badge badge-info">Alineado (${formatPct(pct, 1)})</span>`;
        if (pct <= -5) badge = `<span class="badge badge-success"><i class="fa-solid fa-trophy"></i> Ventaja (${formatPct(pct, 1)})</span>`;
        else if (pct >= 8) badge = `<span class="badge badge-warning">Sobre Moda (+${formatPct(pct, 1)})</span>`;

        return `
            <tr>
                <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(r.frumusaCod)}</code></td>
                <td style="font-weight:700;">${escapeHtml(r.frumusaName)}</td>
                <td><span class="badge badge-neutral">${escapeHtml(r.frumusaUnit)}</span></td>
                <td class="num-col" style="font-weight:800;">${formatCRC(liveInterfrutdPrice)}</td>
                <td>${escapeHtml(r.cenadaName)}</td>
                <td class="num-col" style="font-weight:700;">${formatCRC(cenadaPrice)}</td>
                <td class="num-col" style="color:${diff <= 0 ? 'var(--success)' : 'var(--warning)'}; font-weight:700;">
                    ${diff > 0 ? '+' : ''}${formatCRC(diff)}
                </td>
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
    if (clientSel) {
        const curVal = clientSel.value || appState.activeQuoteDraft?.clienteId;
        clientSel.innerHTML = appState.clients.map(c =>
            `<option value="${escapeHtml(c.id)}">${escapeHtml(c.nombre)} (${escapeHtml(c.segmento)})</option>`
        ).join('');
        if (curVal && appState.clients.some(c => c.id === curVal)) {
            clientSel.value = curVal;
        }
    }

    const skuSel = document.getElementById('quoteAddSkuSelect');
    if (skuSel) {
        const currentSelectedSku = skuSel.value;
        const activeList = appState.activeQuoteDraft?.listaPrecio || 'HORECA';
        const sorted = [...appState.catalog].sort((a, b) => a.name.localeCompare(b.name));
        skuSel.innerHTML = sorted.map(item => {
            const p = getBasePriceByList(item, activeList);
            return `<option value="${escapeHtml(item.sku)}">[${escapeHtml(item.sku)}] ${escapeHtml(item.name)} (${escapeHtml(item.unidad)}) — ${activeList}: ${formatCRC(p)}</option>`;
        }).join('');
        if (currentSelectedSku) skuSel.value = currentSelectedSku;
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
    if (titleEl) titleEl.textContent = `Cotizador Comercial B2B (${draft.id})`;

    const clientSel = document.getElementById('quoteClientSelect');
    if (clientSel) clientSel.value = draft.clienteId;

    const listSel = document.getElementById('quotePriceListSelect');
    if (listSel) listSel.value = draft.listaPrecio;

    const validIn = document.getElementById('quoteValidUntil');
    if (validIn) validIn.value = draft.vigenciaHasta;

    const notesIn = document.getElementById('quoteNotesInput');
    if (notesIn) notesIn.value = draft.notas || '';

    // Calcular totales de la cotización en edición (8 columnas limpias, modelo contra pedido)
    let draftRevenue = 0;
    let draftCost = 0;
    const tbody = document.getElementById('quoteLinesTableBody');

    if (tbody) {
        tbody.innerHTML = draft.items.map((line, idx) => {
            const precioFinal = Math.round(line.precioLista * (1 - (Number(line.descuentoPct) || 0) / 100));
            line.precioFinal = precioFinal;
            const subtotal = line.cantidad * precioFinal;
            const subtotalCost = line.cantidad * line.costoUnit;
            draftRevenue += subtotal;
            draftCost += subtotalCost;
            const lineMarginPct = precioFinal > 0 ? ((precioFinal - line.costoUnit) / precioFinal) * 100 : 0;
            const marginClass = lineMarginPct < 16 ? 'badge-danger' : (lineMarginPct < 22 ? 'badge-warning' : 'badge-success');

            return `
                <tr>
                    <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(line.sku)}</code></td>
                    <td>
                        <div style="font-weight:600;">${escapeHtml(line.name)}</div>
                        <div style="font-size:11px; color:var(--text-muted);">Und: ${escapeHtml(line.unidad)} · Lista: ${formatCRC(line.precioLista)}</div>
                    </td>
                    <td class="num-col">
                        <input type="number" class="inline-price-input quote-line-qty" data-idx="${idx}" value="${line.cantidad}" min="1" style="width:68px;" aria-label="Cantidad para ${escapeHtml(line.name)}">
                    </td>
                    <td class="num-col">${formatCRC(line.costoUnit)}</td>
                    <td class="num-col">
                        <input type="number" class="inline-price-input quote-line-disc" data-idx="${idx}" value="${line.descuentoPct}" min="0" max="35" step="0.5" style="width:60px;" aria-label="Porcentaje de descuento para ${escapeHtml(line.name)}">
                    </td>
                    <td class="num-col" style="font-weight:800;">${formatCRC(precioFinal)}</td>
                    <td class="num-col" style="font-weight:700;">${formatCRC(subtotal)}</td>
                    <td class="center-col" style="white-space:nowrap;">
                        <div style="display:inline-flex; align-items:center; gap:6px;">
                            <span class="badge ${marginClass}">${formatPct(lineMarginPct, 1)}</span>
                            <button type="button" class="btn btn-danger btn-sm btn-remove-qline" data-idx="${idx}" title="Eliminar línea">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </div>
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
        if (appState.quotes.length === 0) {
            listCont.innerHTML = `
                <div style="background:var(--bg-elevated); border:1px dashed var(--border-color); border-radius:var(--radius-sm); padding:24px 16px; text-align:center; color:var(--text-muted); font-size:13px;">
                    <i class="fa-solid fa-folder-open" style="font-size:24px; color:var(--text-muted); margin-bottom:8px; display:block;"></i>
                    No hay cotizaciones guardadas en cartera activa.
                </div>
            `;
        } else {
            listCont.innerHTML = appState.quotes.map(q => {
                const rev = q.items.reduce((s, i) => s + i.cantidad * (i.precioFinal || i.precioLista), 0);
                const cost = q.items.reduce((s, i) => s + i.cantidad * i.costoUnit, 0);
                const marg = rev > 0 ? ((rev - cost) / rev) * 100 : 0;
                const isCurrent = q.id === draft.id;

                let statusBadge = `<span class="badge badge-info">${escapeHtml(q.estado)}</span>`;
                if (q.estado === 'APROBADA' || q.estado === 'RECIBIDA_CLIENTE') {
                    statusBadge = `<span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> Recibida / Aprobada</span>`;
                } else if (q.estado === 'CONVERTIDA_OC') {
                    statusBadge = `<span class="badge badge-success"><i class="fa-solid fa-truck-fast"></i> Convertida OC</span>`;
                }

                const isApprovedOrOc = q.estado === 'APROBADA' || q.estado === 'RECIBIDA_CLIENTE' || q.estado === 'CONVERTIDA_OC';

                return `
                    <div style="background:var(--bg-elevated); border:1px solid ${isCurrent ? 'var(--brand-green)' : 'var(--border-color)'}; border-radius:var(--radius-sm); padding:13px 15px;">
                        <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                            <div>
                                <strong style="color:var(--brand-green); font-size:14px;">${escapeHtml(q.id)}</strong>
                                <span style="font-size:12px; color:var(--text-muted); margin-left:6px;">${escapeHtml(q.fecha)}</span>
                                <div style="font-weight:700; font-size:14px; margin-top:2px;">${escapeHtml(q.clienteNombre)}</div>
                            </div>
                            ${statusBadge}
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center; font-size:12px; color:var(--text-secondary); margin:8px 0;">
                            <span>Lista: <strong>${escapeHtml(q.listaPrecio)}</strong> · ${q.items.length} SKUs</span>
                            <span>Monto: <strong style="color:var(--text-main);">${formatCRC(rev)}</strong> (Margen: <strong>${formatPct(marg, 1)}</strong>)</span>
                        </div>
                        <div style="display:flex; gap:6px; flex-wrap:wrap; margin-top:10px; align-items:center;">
                            <button type="button" class="btn btn-secondary btn-sm btn-load-quote" data-qid="${escapeHtml(q.id)}" title="Cargar en el constructor para editar">
                                <i class="fa-solid fa-pen-to-square"></i> Editar
                            </button>
                            ${!isApprovedOrOc ? `
                            <button type="button" class="btn btn-secondary btn-sm btn-approve-quote" data-qid="${escapeHtml(q.id)}" title="Marcar como Recibida / Aprobada por el cliente">
                                <i class="fa-solid fa-clipboard-check" style="color:var(--brand-green-light);"></i> Dar Recibida
                            </button>` : ''}
                            <button type="button" class="btn btn-primary btn-sm btn-convert-quote-oc" data-qid="${escapeHtml(q.id)}" title="Generar Orden de Compra (OC)">
                                <i class="fa-solid fa-truck-Arrow-right"></i> Convertir a OC
                            </button>
                            <button type="button" class="btn btn-danger btn-sm btn-delete-quote" data-qid="${escapeHtml(q.id)}" title="Eliminar cotización de cartera" style="margin-left:auto;">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </div>
                    </div>
                `;
            }).join('');
        }
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
                    <p style="margin:2px 0 0; font-size:13px; color:#475569;">Departamento Comercial & Pricing Corporativo · Santo Domingo, Heredia, Costa Rica</p>
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

            <table style="width:100%; border-collapse:collapse; font-size:13px; margin-bottom:16px;">
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
function renderOcStepper(status) {
    if (status === 'ALERTA_STOCK') {
        return `<span class="badge badge-danger"><i class="fa-solid fa-triangle-exclamation"></i> Alerta Stock Parcial</span>`;
    }
    const steps = [
        { key: 'OC_RECIBIDA', label: '1. Recibida' },
        { key: 'EN_ALISTAMIENTO', label: '2. Alistamiento' },
        { key: 'EN_RUTA', label: '3. En Ruta' },
        { key: 'ENTREGADA', label: '4. Entregada' }
    ];
    const curIdx = steps.findIndex(s => s.key === status);
    return `
        <div class="oc-stepper">
            ${steps.map((st, i) => {
                let cls = 'oc-step';
                if (i === curIdx) cls += ' active';
                else if (curIdx !== -1 && i < curIdx) cls += ' completed';
                return `<span class="${cls}">${st.label}</span>`;
            }).join('')}
        </div>
    `;
}

function openOcDetailModal(oc) {
    const modal = document.getElementById('ocDetailModal');
    const title = document.getElementById('ocDetailModalTitle');
    const cont = document.getElementById('ocDetailModalContent');
    if (!modal || !cont) return;

    const totalPed = oc.lineas.reduce((s, l) => s + l.cantPedida, 0);
    const totalDesp = oc.lineas.reduce((s, l) => s + l.cantDespachada, 0);
    const totalMonto = oc.lineas.reduce((s, l) => s + (l.cantDespachada * l.precioUnit), 0);
    const totalCosto = oc.lineas.reduce((s, l) => s + (l.cantDespachada * l.costoUnit), 0);
    const totalMargen = totalMonto > 0 ? ((totalMonto - totalCosto) / totalMonto) * 100 : 0;
    const overallFill = totalPed > 0 ? (totalDesp / totalPed) * 100 : 100;
    const isSuper = oc.segmento.toLowerCase().includes('supermercado');
    const slaTarget = isSuper ? 98 : 95;
    const cumpleSla = overallFill >= slaTarget;

    if (title) title.textContent = `Auditoría OTIF — Orden ${oc.ocNumero} (${oc.clienteNombre})`;
    cont.innerHTML = `
        <div style="display:flex; justify-content:space-between; flex-wrap:wrap; gap:12px; background:var(--bg-elevated); padding:12px 16px; border-radius:var(--radius-sm); margin-bottom:16px; border:1px solid var(--border-color);">
            <div><span style="color:var(--text-muted); font-size:11px; text-transform:uppercase;">Segmento / Canal</span><div style="font-weight:700;">${escapeHtml(oc.segmento)} · ${escapeHtml(oc.canal)}</div></div>
            <div><span style="color:var(--text-muted); font-size:11px; text-transform:uppercase;">Entrega Compromiso</span><div style="font-weight:700;">${escapeHtml(oc.fechaEntregaCompromiso)}</div></div>
            <div><span style="color:var(--text-muted); font-size:11px; text-transform:uppercase;">Fill Rate / SLA (${slaTarget}%)</span><div style="font-weight:800; font-size:15px; color:${cumpleSla ? 'var(--brand-lime)' : 'var(--danger)'};">${formatPct(overallFill, 1)} ${cumpleSla ? '✅' : '⚠️'}</div></div>
            <div><span style="color:var(--text-muted); font-size:11px; text-transform:uppercase;">Facturación Efectiva</span><div style="font-weight:800; font-size:15px; color:var(--text-main);">${formatCRC(totalMonto)} (Margen: ${formatPct(totalMargen, 1)})</div></div>
        </div>
        <div class="table-scroll">
            <table class="data-table" style="font-size:12px;">
                <thead>
                    <tr>
                        <th>SKU</th>
                        <th>Producto Solicitado</th>
                        <th class="num-col">Pedida</th>
                        <th class="num-col">Despachada</th>
                        <th class="num-col">Cumplimiento %</th>
                        <th class="num-col">Precio Unit (₡)</th>
                        <th class="num-col">Subtotal Facturado</th>
                        <th class="center-col">Auditoría de Despacho</th>
                    </tr>
                </thead>
                <tbody>
                    ${oc.lineas.map(l => {
                        const lFill = l.cantPedida > 0 ? (l.cantDespachada / l.cantPedida) * 100 : 100;
                        const isFull = l.cantDespachada >= l.cantPedida;
                        const diff = l.cantPedida - l.cantDespachada;
                        return `
                            <tr>
                                <td><code style="font-weight:700; color:var(--brand-green);">${escapeHtml(l.sku)}</code></td>
                                <td style="font-weight:700;">${escapeHtml(l.name)}</td>
                                <td class="num-col">${formatNum(l.cantPedida)}</td>
                                <td class="num-col" style="font-weight:700;">${formatNum(l.cantDespachada)}</td>
                                <td class="num-col">
                                    <span class="badge ${lFill >= 98 ? 'badge-success' : (lFill >= 90 ? 'badge-warning' : 'badge-danger')}">
                                        ${formatPct(lFill, 1)}
                                    </span>
                                </td>
                                <td class="num-col">${formatCRC(l.precioUnit)}</td>
                                <td class="num-col" style="font-weight:700;">${formatCRC(l.cantDespachada * l.precioUnit)}</td>
                                <td class="center-col">
                                    ${isFull
                                        ? '<span class="badge badge-success"><i class="fa-solid fa-check-double"></i> 100% Despachado</span>'
                                        : `<span class="badge badge-danger"><i class="fa-solid fa-triangle-exclamation"></i> Faltante: ${formatNum(diff)} und</span>`
                                    }
                                </td>
                            </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
        </div>
    `;
    modal.classList.add('open');
}

function openClientModal(clientId = null) {
    const modal = document.getElementById('clientFormModal');
    const title = document.getElementById('clientModalTitle');
    const formId = document.getElementById('clientFormId');
    const nameIn = document.getElementById('clientFormName');
    const segIn = document.getElementById('clientFormSegment');
    const listIn = document.getElementById('clientFormList');
    const contactIn = document.getElementById('clientFormContact');
    const emailIn = document.getElementById('clientFormEmail');
    const phoneIn = document.getElementById('clientFormPhone');
    const creditIn = document.getElementById('clientFormCreditDays');
    const routeIn = document.getElementById('clientFormRoute');
    const dockIn = document.getElementById('clientFormDockWindow');
    const cutOffIn = document.getElementById('clientFormCutOff');
    const slaIn = document.getElementById('clientFormSlaMeta');
    const dayPills = document.querySelectorAll('#clientFormDayPills .day-pill-btn');

    if (!modal) return;

    if (clientId) {
        const cli = appState.clients.find(c => c.id === clientId);
        if (!cli) return;
        if (title) title.innerHTML = `<i class="fa-solid fa-pen-to-square" style="color:var(--brand-green-light); margin-right:6px;"></i>Editar Ficha — ${escapeHtml(cli.nombre)} (${cli.id})`;
        if (formId) formId.value = cli.id;
        if (nameIn) nameIn.value = cli.nombre || '';
        if (segIn) segIn.value = cli.segmento || 'HORECA Premium';
        if (listIn) listIn.value = cli.listaAsignada || 'HORECA';
        if (contactIn) contactIn.value = cli.contacto || '';
        if (emailIn) emailIn.value = cli.email || '';
        if (phoneIn) phoneIn.value = cli.telefono || '';
        if (creditIn) creditIn.value = cli.diasCredito !== undefined ? cli.diasCredito : 15;
        if (routeIn) routeIn.value = cli.ruta || 'Ruta GAM Central (402)';
        if (dockIn) dockIn.value = cli.ventanaAnden || '05:30 AM – 07:00 AM';
        if (cutOffIn) cutOffIn.value = cli.horaCorte || 'Día ant. 03:00 PM';
        if (slaIn) slaIn.value = cli.slaMeta !== undefined ? cli.slaMeta : 95;

        const activeDays = cli.diasEntrega || ['L', 'K', 'V'];
        dayPills.forEach(btn => {
            const d = btn.getAttribute('data-formday');
            const isActive = activeDays.includes(d);
            btn.classList.toggle('active', isActive);
            btn.setAttribute('aria-pressed', isActive ? 'true' : 'false');
        });
    } else {
        if (title) title.innerHTML = `<i class="fa-solid fa-user-plus" style="color:var(--brand-green-light); margin-right:6px;"></i>Registrar Nuevo Cliente Corporativo B2B`;
        if (formId) formId.value = '';
        if (nameIn) nameIn.value = '';
        if (segIn) segIn.value = 'HORECA Premium';
        if (listIn) listIn.value = 'HORECA';
        if (contactIn) contactIn.value = '';
        if (emailIn) emailIn.value = '';
        if (phoneIn) phoneIn.value = '';
        if (creditIn) creditIn.value = 15;
        if (routeIn) routeIn.value = 'Ruta GAM Central (402)';
        if (dockIn) dockIn.value = '05:30 AM – 07:00 AM';
        if (cutOffIn) cutOffIn.value = 'Día ant. 03:00 PM';
        if (slaIn) slaIn.value = 95;

        const defaultDays = ['L', 'K', 'V'];
        dayPills.forEach(btn => {
            const d = btn.getAttribute('data-formday');
            const isActive = defaultDays.includes(d);
            btn.classList.toggle('active', isActive);
            btn.setAttribute('aria-pressed', isActive ? 'true' : 'false');
        });
    }

    modal.classList.add('open');
    setTimeout(() => nameIn?.focus(), 60);
}

function saveClientFormData() {
    const formId = document.getElementById('clientFormId')?.value.trim();
    const nameVal = document.getElementById('clientFormName')?.value.trim();
    if (!nameVal) {
        showToast('El nombre o razón social del cliente es obligatorio.', 'fa-triangle-exclamation');
        document.getElementById('clientFormName')?.focus();
        return;
    }

    const segVal = document.getElementById('clientFormSegment')?.value || 'HORECA Premium';
    const listVal = document.getElementById('clientFormList')?.value || 'HORECA';
    const contactVal = document.getElementById('clientFormContact')?.value.trim() || 'Compras Corporativas';
    const emailVal = document.getElementById('clientFormEmail')?.value.trim() || 'compras@cliente.com';
    const phoneVal = document.getElementById('clientFormPhone')?.value.trim() || '2200-0000';
    const creditVal = Math.max(0, parseInt(document.getElementById('clientFormCreditDays')?.value, 10) || 0);
    const routeVal = document.getElementById('clientFormRoute')?.value.trim() || 'Ruta GAM Central (402)';
    const dockVal = document.getElementById('clientFormDockWindow')?.value.trim() || '05:30 AM – 07:00 AM';
    const cutOffVal = document.getElementById('clientFormCutOff')?.value.trim() || 'Día ant. 03:00 PM';
    const slaVal = Math.min(100, Math.max(80, parseInt(document.getElementById('clientFormSlaMeta')?.value, 10) || 95));

    const selectedDays = [];
    document.querySelectorAll('#clientFormDayPills .day-pill-btn.active').forEach(b => {
        selectedDays.push(b.getAttribute('data-formday'));
    });
    if (selectedDays.length === 0) {
        selectedDays.push('L');
    }

    if (formId) {
        const cli = appState.clients.find(c => c.id === formId);
        if (cli) {
            cli.nombre = nameVal;
            cli.segmento = segVal;
            cli.listaAsignada = listVal;
            cli.contacto = contactVal;
            cli.email = emailVal;
            cli.telefono = phoneVal;
            cli.diasCredito = creditVal;
            cli.ruta = routeVal;
            cli.ventanaAnden = dockVal;
            cli.horaCorte = cutOffVal;
            cli.slaMeta = slaVal;
            cli.diasEntrega = selectedDays;

            const baseCli = B2B_CLIENTS.find(c => c.id === formId);
            if (baseCli) Object.assign(baseCli, cli);

            showToast(`Ficha de ${cli.nombre} actualizada exitosamente.`, 'fa-circle-check');
        }
    } else {
        let maxNum = 0;
        appState.clients.forEach(c => {
            const m = String(c.id).match(/\d+/);
            if (m) {
                const n = parseInt(m[0], 10);
                if (n > maxNum) maxNum = n;
            }
        });
        const nextId = `CLI-00${maxNum + 1}`;
        const newClient = {
            id: nextId,
            nombre: nameVal,
            contacto: contactVal,
            email: emailVal,
            telefono: phoneVal,
            segmento: segVal,
            listaAsignada: listVal,
            descuentoBasePct: 0,
            diasCredito: creditVal,
            ruta: routeVal,
            ventanaAnden: dockVal,
            horaCorte: cutOffVal,
            slaMeta: slaVal,
            diasEntrega: selectedDays
        };
        appState.clients.push(newClient);
        showToast(`Nuevo cliente ${newClient.nombre} (${newClient.id}) creado en cartera.`, 'fa-user-check');
    }

    try {
        localStorage.setItem('interfrutd_clients_v3', JSON.stringify(appState.clients));
        const daysMap = {};
        appState.clients.forEach(c => { daysMap[c.id] = c.diasEntrega; });
        localStorage.setItem('interfrutd_client_days', JSON.stringify(daysMap));
    } catch (e) {}

    document.getElementById('clientFormModal')?.classList.remove('open');
    populateQuoteSelectors();
    renderOrdersAndInventoryModule();
    if (appState.activeTab === 'cotizaciones') renderQuotesModule();
}

function renderOrdersAndInventoryModule() {
    const baseCatalog = getFilteredCatalog();
    const totalMonthlySales = baseCatalog.reduce((s, x) => s + x.ventasMensualesCrc, 0);

    const dayNamesMap = {
        ALL: 'Semana Completa (Lunes a Sábado)',
        L: 'Lunes',
        M: 'Martes',
        K: 'Miércoles',
        J: 'Jueves',
        V: 'Viernes',
        S: 'Sábado'
    };
    const activeDay = appState.selectedScheduleDay || 'ALL';
    const q = appState.searchQuery.trim().toLowerCase();

    // Filtrar clientes según el Día de Despacho seleccionado y el buscador global
    const filteredClients = appState.clients.filter(cli => {
        if (activeDay !== 'ALL' && !(cli.diasEntrega || []).includes(activeDay)) return false;
        if (q) {
            const match = cli.nombre.toLowerCase().includes(q) ||
                cli.segmento.toLowerCase().includes(q) ||
                (cli.ruta && cli.ruta.toLowerCase().includes(q)) ||
                (cli.contacto && cli.contacto.toLowerCase().includes(q));
            if (!match) return false;
        }
        return true;
    });

    // 4 Canales Principales para el Monitor OTIF (Operación Contra Pedido)
    const channelSegments = [
        { key: 'Supermercados', label: 'Supermercados', icon: 'fa-store', sla: 98, hint: 'AutoMercado / Walmart' },
        { key: 'HORECA Premium', label: 'HORECA Premium', icon: 'fa-hotel', sla: 95, hint: 'Hoteles / Restaurantes' },
        { key: 'Cadenas de Restaurantes', label: 'Cadenas Gastronómicas', icon: 'fa-utensils', sla: 95, hint: 'Rostipollos & Casual' },
        { key: 'Institucional / Catering', label: 'Comedores / Catering', icon: 'fa-building-user', sla: 95, hint: 'Sodexo Zona Franca' }
    ];

    const segStats = {};
    channelSegments.forEach(c => {
        segStats[c.key] = { pedidas: 0, despachadas: 0, facturado: 0, count: 0, sla: c.sla };
    });

    let totalOrderedUnits = 0;
    let totalDispatchedUnits = 0;
    let totalOcFacturado = 0;
    let totalOcPedidoPotencial = 0;
    let ocs100Count = 0;

    const dayFilteredOCs = appState.purchaseOrders.filter(oc => {
        if (activeDay !== 'ALL' && oc.diaSemana && oc.diaSemana !== activeDay) return false;
        return true;
    });

    dayFilteredOCs.forEach(oc => {
        let ocPed = 0;
        let ocDesp = 0;
        let ocFact = 0;
        let ocPot = 0;

        oc.lineas.forEach(l => {
            ocPed += l.cantPedida;
            ocDesp += l.cantDespachada;
            ocFact += l.cantDespachada * l.precioUnit;
            ocPot += l.cantPedida * l.precioUnit;
        });

        totalOrderedUnits += ocPed;
        totalDispatchedUnits += ocDesp;
        totalOcFacturado += ocFact;
        totalOcPedidoPotencial += ocPot;
        if (ocDesp >= ocPed && ocPed > 0) ocs100Count++;

        if (segStats[oc.segmento]) {
            segStats[oc.segmento].pedidas += ocPed;
            segStats[oc.segmento].despachadas += ocDesp;
            segStats[oc.segmento].facturado += ocFact;
            segStats[oc.segmento].count += 1;
        }
    });

    const globalFillRatePct = totalOrderedUnits > 0 ? (totalDispatchedUnits / totalOrderedUnits) * 100 : 100;

    // 1. Grid Principal de 4 KPIs Ejecutivos
    const kpiGrid = document.getElementById('ordersKpiGrid');
    if (kpiGrid) {
        kpiGrid.innerHTML = `
            <div class="kpi-card">
                <div class="kpi-header"><span>Clientes Programados (${activeDay === 'ALL' ? 'Semana' : dayNamesMap[activeDay]})</span><i class="fa-solid fa-calendar-check"></i></div>
                <div class="kpi-value">${filteredClients.length} <span style="font-size:14px; font-weight:600; color:var(--text-muted);">/ ${appState.clients.length} B2B</span></div>
                <div class="kpi-sub">
                    <span class="badge badge-success">Rutas Matutinas 4:00–8:00 AM</span>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-header"><span>Facturación Contra Pedido</span><i class="fa-solid fa-file-invoice-dollar"></i></div>
                <div class="kpi-value">${formatCRC(totalOcFacturado)}</div>
                <div class="kpi-sub">
                    <span class="badge badge-success">${dayFilteredOCs.length} OCs</span>
                    <span>Proyección Mes: <strong>${formatCRC(totalMonthlySales)}</strong></span>
                </div>
            </div>
            <div class="kpi-card ${globalFillRatePct >= 96 ? '' : 'kpi-warning'}">
                <div class="kpi-header"><span>Nivel de Servicio (Fill Rate OTIF)</span><i class="fa-solid fa-truck-fast"></i></div>
                <div class="kpi-value">${formatPct(globalFillRatePct, 1)}</div>
                <div class="kpi-sub">
                    <span class="badge badge-success">${ocs100Count} de ${dayFilteredOCs.length} al 100%</span>
                    <span>${formatNum(totalDispatchedUnits)} / ${formatNum(totalOrderedUnits)} und</span>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-header"><span>SLA Cadenas Supermercados</span><i class="fa-solid fa-store"></i></div>
                <div class="kpi-value">${formatPct(segStats['Supermercados']?.pedidas > 0 ? (segStats['Supermercados'].despachadas / segStats['Supermercados'].pedidas) * 100 : 99.4, 1)}</div>
                <div class="kpi-sub">
                    <span class="badge badge-success"><i class="fa-solid fa-shield-check"></i> Meta &ge;98.0%</span>
                    <span>Cero merma en bodega</span>
                </div>
            </div>
        `;
    }

    // 2. Renderizar Matriz de Programación de Clientes B2B, Días de Entrega y Rutas
    const headerTitleEl = document.getElementById('clientScheduleHeaderTitle');
    if (headerTitleEl) {
        headerTitleEl.textContent = `Programación de Clientes B2B y Rutas — ${dayNamesMap[activeDay]}`;
    }

    const clientTbody = document.getElementById('clientScheduleTableBody');
    if (clientTbody) {
        const allDays = [
            { code: 'L', title: 'Lunes' },
            { code: 'M', title: 'Martes' },
            { code: 'K', title: 'Miércoles' },
            { code: 'J', title: 'Jueves' },
            { code: 'V', title: 'Viernes' },
            { code: 'S', title: 'Sábado' }
        ];

        clientTbody.innerHTML = filteredClients.map(cli => {
            const dias = cli.diasEntrega || [];
            const dayButtonsHtml = allDays.map(d => {
                const isOn = dias.includes(d.code);
                return `<button type="button" class="day-pill-btn ${isOn ? 'active' : ''}" data-cli="${escapeHtml(cli.id)}" data-daycode="${d.code}" aria-pressed="${isOn ? 'true' : 'false'}" aria-label="Entrega ${d.title} para ${escapeHtml(cli.nombre)}: ${isOn ? 'Activa' : 'Inactiva'}" title="${d.title}: ${isOn ? 'Entrega Activa' : 'Sin Entrega'}">${d.code}</button>`;
            }).join('');

            return `
                <tr>
                    <td>
                        <div style="font-weight:700;">${escapeHtml(cli.nombre)}</div>
                        <div style="font-size:11px; color:var(--text-muted);">${escapeHtml(cli.segmento)} · ${escapeHtml(cli.contacto)}</div>
                    </td>
                    <td>
                        <div style="font-weight:600; color:var(--text-main);"><i class="fa-solid fa-route" style="color:var(--brand-green-light); margin-right:4px;"></i>${escapeHtml(cli.ruta)}</div>
                        <div style="font-size:11px; color:var(--brand-lime);">Andén: ${escapeHtml(cli.ventanaAnden || '05:00 AM – 07:00 AM')}</div>
                    </td>
                    <td style="white-space:nowrap;">
                        <span class="badge badge-neutral"><i class="fa-regular fa-clock"></i> ${escapeHtml(cli.horaCorte || 'Día ant. 02:00 PM')}</span>
                    </td>
                    <td class="center-col" style="white-space:nowrap;">
                        <div class="day-pills-wrap">${dayButtonsHtml}</div>
                        <div style="font-size:10px; color:var(--text-muted); margin-top:2px;">${dias.length} recepciones/sem</div>
                    </td>
                    <td style="white-space:nowrap;">
                        <span class="badge badge-info">${escapeHtml(cli.listaAsignada)} (${cli.descuentoBasePct > 0 ? '-' + cli.descuentoBasePct + '%' : 'Base'})</span>
                        <span style="font-size:11px; color:var(--text-muted); margin-left:6px;">${cli.diasCredito} días</span>
                    </td>
                    <td class="center-col" style="white-space:nowrap;">
                        <div style="display:inline-flex; align-items:center; gap:5px;">
                            <span class="badge badge-success">SLA &ge;${cli.slaMeta || 95}%</span>
                            <button type="button" class="btn btn-secondary btn-sm btn-edit-client" data-cli="${escapeHtml(cli.id)}" title="Editar Ficha de Cliente">
                                <i class="fa-solid fa-pen-to-square"></i>
                            </button>
                            <button type="button" class="btn btn-secondary btn-sm btn-schedule-quote" data-cli="${escapeHtml(cli.id)}" title="Crear Cotización / Orden para este cliente">
                                <i class="fa-solid fa-file-circle-plus"></i> Cotizar
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
    }

    // 3. Renderizar Monitor OTIF por Canal (4 tarjetas en 1 fila)
    const otifGridEl = document.getElementById('otifChannelGrid');
    if (otifGridEl) {
        otifGridEl.innerHTML = channelSegments.map(c => {
            const stat = segStats[c.key] || { pedidas: 0, despachadas: 0, facturado: 0, count: 0, sla: c.sla };
            const rate = stat.pedidas > 0 ? (stat.despachadas / stat.pedidas) * 100 : 100;
            const meets = rate >= c.sla;
            const barColor = meets ? 'var(--brand-green-light)' : 'var(--danger)';

            return `
                <div class="otif-card">
                    <div class="otif-card-top">
                        <span class="otif-card-name"><i class="fa-solid ${c.icon}"></i> ${c.label}</span>
                        <span class="otif-card-sla">SLA &ge;${c.sla}%</span>
                    </div>
                    <div class="otif-card-metric">
                        <div class="otif-val" style="color:${meets ? 'var(--text-main)' : 'var(--danger)'};">${formatPct(rate, 1)}</div>
                        <span class="badge ${meets ? 'badge-success' : 'badge-danger'}">
                            ${meets ? 'Cumple' : 'Riesgo'}
                        </span>
                    </div>
                    <div class="otif-progress-wrap">
                        <div class="otif-progress-bar" style="width:${Math.min(100, rate)}%; background:${barColor};"></div>
                    </div>
                    <div class="otif-card-footer">
                        <span>${c.hint}</span>
                        <span>${formatCRC(stat.facturado)}</span>
                    </div>
                </div>
            `;
        }).join('');
    }

    const badgeOC = document.getElementById('badgeCountOC');
    if (badgeOC) badgeOC.textContent = String(appState.clients.length);

    // 4. Tabla Simplificada de Pipeline de Órdenes de Compra (8 Columnas Clave)
    const ocTbody = document.getElementById('purchaseOrdersTableBody');
    if (ocTbody) {
        const filteredOCs = dayFilteredOCs.filter(oc => {
            const ped = oc.lineas.reduce((s, l) => s + l.cantPedida, 0);
            const desp = oc.lineas.reduce((s, l) => s + l.cantDespachada, 0);
            if (appState.ocFilterStatus === '100_OTIF') return desp >= ped;
            if (appState.ocFilterStatus === 'PARCIAL') return desp < ped;
            if (appState.ocFilterStatus !== 'ALL') return oc.estado === appState.ocFilterStatus;
            return true;
        });

        if (filteredOCs.length === 0) {
            ocTbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding:22px 14px; color:var(--text-muted); font-size:13px;">No hay órdenes de compra registradas con el filtro activo.</td></tr>`;
            return;
        }

        ocTbody.innerHTML = filteredOCs.map(oc => {
            const pedida = oc.lineas.reduce((s, l) => s + l.cantPedida, 0);
            const desp = oc.lineas.reduce((s, l) => s + l.cantDespachada, 0);
            const monto = oc.lineas.reduce((s, l) => s + (l.cantDespachada * l.precioUnit), 0);
            const costo = oc.lineas.reduce((s, l) => s + (l.cantDespachada * l.costoUnit), 0);
            const margenOc = monto > 0 ? ((monto - costo) / monto) * 100 : 0;
            const fillRate = pedida > 0 ? (desp / pedida) * 100 : 100;

            const isSuper = oc.segmento.toLowerCase().includes('supermercado');
            const slaMeta = isSuper ? 98 : 95;
            const cumpleSla = fillRate >= slaMeta;

            return `
                <tr>
                    <td style="white-space:nowrap;">
                        <code style="font-weight:800; color:var(--brand-lime);">${escapeHtml(oc.ocNumero)}</code>
                        <div style="font-size:11px; color:var(--text-muted);">Ref: ${escapeHtml(oc.cotizacionRef)}</div>
                    </td>
                    <td>
                        <div style="font-weight:700;">${escapeHtml(oc.clienteNombre)}</div>
                        <div style="font-size:11px; color:var(--text-muted);">${escapeHtml(oc.segmento)} · ${oc.lineas.length} SKUs</div>
                    </td>
                    <td style="white-space:nowrap;">
                        <div style="font-weight:600;">${escapeHtml(oc.fechaEntregaCompromiso)}</div>
                        <div style="font-size:11px; color:var(--text-muted);">${escapeHtml(oc.ventanaAnden || 'Andén Matutino')}</div>
                    </td>
                    <td class="num-col" style="white-space:nowrap;">
                        <strong>${formatNum(desp)}</strong> <span style="color:var(--text-muted);">/ ${formatNum(pedida)} und</span>
                    </td>
                    <td class="center-col" style="white-space:nowrap;">
                        <span class="badge ${cumpleSla ? 'badge-success' : 'badge-danger'}">
                            ${formatPct(fillRate, 1)} · SLA &ge;${slaMeta}%
                        </span>
                    </td>
                    <td class="num-col" style="white-space:nowrap;">
                        <div style="font-weight:700;">${formatCRC(monto)}</div>
                        <div style="font-size:11px; color:var(--brand-lime);">Margen: ${formatPct(margenOc, 1)}</div>
                    </td>
                    <td>${renderOcStepper(oc.estado)}</td>
                    <td class="center-col" style="white-space:nowrap;">
                        <div style="display:inline-flex; align-items:center; gap:5px;">
                            <button type="button" class="btn btn-secondary btn-sm btn-view-oc" data-oc="${escapeHtml(oc.ocNumero)}" title="Auditar líneas de la orden">
                                <i class="fa-solid fa-list-check"></i>
                            </button>
                            ${oc.estado !== 'ENTREGADA' ? `
                            <button type="button" class="btn btn-secondary btn-sm btn-mark-delivered-oc" data-oc="${escapeHtml(oc.ocNumero)}" title="Marcar como Recibida / Entregada en Andén en 1 clic">
                                <i class="fa-solid fa-circle-check" style="color:var(--brand-green-light);"></i>
                            </button>
                            <button type="button" class="btn btn-primary btn-sm btn-advance-oc" data-oc="${escapeHtml(oc.ocNumero)}" title="Avanzar etapa logística">
                                <i class="fa-solid fa-forward-step"></i>
                            </button>` : `
                            <span class="badge badge-success" style="font-size:11px;"><i class="fa-solid fa-circle-check"></i> Recibida</span>
                            `}
                            <button type="button" class="btn btn-danger btn-sm btn-delete-oc" data-oc="${escapeHtml(oc.ocNumero)}" title="Eliminar / Cancelar Orden de Compra">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
    }
}

function renderCommercialCharts(catalogList) {
    if (typeof Chart === 'undefined') return;

    // Agrupar ventas y margen por Categoría
    const byCat = {};
    catalogList.forEach(item => {
        if (!byCat[item.categoria]) {
            byCat[item.categoria] = { sales: 0, profit: 0 };
        }
        byCat[item.categoria].sales += item.ventasMensualesCrc;
        byCat[item.categoria].profit += item.utilidadMensualCrc;
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
}

// ============================================================================
// RENDERIZADO GLOBAL Y EVENTOS DE INTERFAZ
// ============================================================================
function renderActiveModule() {
    renderPricingModule();
    if (appState.activeTab === 'cotizaciones') renderQuotesModule();
    if (appState.activeTab === 'ordenes') renderOrdersAndInventoryModule();
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

    // Buscador global + Atajo Ctrl+K (Debounce 120ms para rendimiento a 60fps)
    const searchInput = document.getElementById('globalSearchInput');
    if (searchInput) {
        let searchDebounceTimer = null;
        searchInput.addEventListener('input', () => {
            clearTimeout(searchDebounceTimer);
            searchDebounceTimer = setTimeout(() => {
                appState.searchQuery = searchInput.value;
                renderActiveModule();
            }, 120);
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

    // Política de Márgenes por Tipo de Cadena
    const chainControls = [
        { key: 'SUPERMERCADO', name: 'Supermercados', r: 'rangeMarginSupermercado', n: 'numMarginSupermercado' },
        { key: 'HORECA', name: 'HORECA / Food Service', r: 'rangeMarginHoreca', n: 'numMarginHoreca' },
        { key: 'MAYORISTA', name: 'Mayoristas / CEDI', r: 'rangeMarginMayorista', n: 'numMarginMayorista' },
        { key: 'DETALLE', name: 'Puntos de Venta / Detalle', r: 'rangeMarginDetalle', n: 'numMarginDetalle' }
    ];

    chainControls.forEach(({ key, name, r, n }) => {
        const rEl = document.getElementById(r);
        const nEl = document.getElementById(n);
        if (!rEl || !nEl) return;

        const updateMargin = (val, notify = false) => {
            const num = parseInt(val, 10);
            if (isNaN(num)) return;
            const clamped = Math.max(parseInt(rEl.min, 10), Math.min(parseInt(rEl.max, 10), num));
            appState.channelMargins[key] = clamped;
            rEl.value = clamped;
            nEl.value = clamped;

            try {
                localStorage.setItem('interfrutd_channel_margins', JSON.stringify(appState.channelMargins));
            } catch (e) {}

            appState.catalog.forEach(recalcSkuMetrics);
            renderPricingModule();

            if (notify) {
                showToast(`Margen meta para ${name} ajustado al ${clamped}%`, 'fa-sliders');
            }
        };

        rEl.addEventListener('input', (e) => updateMargin(e.target.value, false));
        rEl.addEventListener('change', (e) => updateMargin(e.target.value, true));
        nEl.addEventListener('change', (e) => updateMargin(e.target.value, true));
    });

    const btnResetMargins = document.getElementById('btnResetChannelMargins');
    if (btnResetMargins) {
        btnResetMargins.addEventListener('click', () => {
            appState.channelMargins = {
                SUPERMERCADO: 24,
                HORECA: 28,
                MAYORISTA: 18,
                DETALLE: 32
            };
            syncChannelMarginInputs();
            try {
                localStorage.removeItem('interfrutd_channel_margins');
                localStorage.removeItem('interfrutd_price_overrides');
                localStorage.removeItem('interfrutd_cost_overrides');
            } catch (e) {}
            appState.catalog = INTERFRUTD_CATALOG.map(item => {
                const copy = { ...item };
                recalcSkuMetrics(copy);
                return copy;
            });
            renderActiveModule();
            showToast('Márgenes corporativos (24% / 28% / 18% / 32%) y precios/costos base restaurados.', 'fa-rotate-left');
        });
    }

    document.querySelectorAll('#pricingQuickChips .chip-btn').forEach(chip => {
        chip.addEventListener('click', () => {
            document.querySelectorAll('#pricingQuickChips .chip-btn').forEach(c => {
                c.classList.remove('active');
                c.setAttribute('aria-pressed', 'false');
            });
            chip.classList.add('active');
            chip.setAttribute('aria-pressed', 'true');
            appState.pricingDiagFilter = chip.getAttribute('data-pfilter');
            renderPricingModule();
        });
    });

    const btnSwitchMatrix = document.getElementById('btnSwitchMatrixView');
    if (btnSwitchMatrix) {
        btnSwitchMatrix.addEventListener('click', () => {
            if (appState.pricingView === 'catalog') {
                appState.pricingView = 'cenada';
                btnSwitchMatrix.innerHTML = `<i class="fa-solid fa-table-list"></i> Volver a Catálogo (${appState.catalog.length})`;
            } else {
                appState.pricingView = 'catalog';
                btnSwitchMatrix.innerHTML = `<i class="fa-solid fa-warehouse"></i> Benchmark CENADA (39)`;
            }
            renderPricingModule();
        });
    }

    // Edición en línea de costos (COGS), precios detalle y botón de enviar a cotización
    const pricingTbody = document.getElementById('pricingTableBody');
    if (pricingTbody) {
        pricingTbody.addEventListener('change', (e) => {
            if (e.target.classList.contains('inline-cost-input')) {
                const sku = e.target.getAttribute('data-sku');
                const newCost = Math.max(5, Number(e.target.value) || 0);
                const item = appState.catalog.find(x => x.sku === sku);
                if (item) {
                    item.costoReposicion = newCost;
                    recalcSkuMetrics(item);
                    try {
                        const raw = localStorage.getItem('interfrutd_cost_overrides');
                        const map = raw ? JSON.parse(raw) : {};
                        map[sku] = newCost;
                        localStorage.setItem('interfrutd_cost_overrides', JSON.stringify(map));
                    } catch (err) {}
                    renderPricingModule();
                    showToast(`Costo COGS actualizado para SKU ${sku}: ${formatCRC(newCost)} (Margen: ${formatPct(item.margenReposicionPct, 1)})`, 'fa-coins');
                }
                return;
            }

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
                showToast(`${count} SKUs ajustados al margen meta (+${formatCRC(monthlyGain)}/mes de utilidad)`, 'fa-sliders');
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

            const btnApprove = e.target.closest('.btn-approve-quote');
            if (btnApprove) {
                const qid = btnApprove.getAttribute('data-qid');
                const quote = appState.quotes.find(q => q.id === qid);
                if (!quote) return;
                quote.estado = 'RECIBIDA_CLIENTE';
                try {
                    localStorage.setItem('interfrutd_quotes', JSON.stringify(appState.quotes));
                } catch (err) {}
                renderQuotesModule();
                showToast(`Cotización ${quote.id} marcada como Recibida / Aprobada por el cliente`, 'fa-circle-check');
                return;
            }

            const btnDelete = e.target.closest('.btn-delete-quote');
            if (btnDelete) {
                const qid = btnDelete.getAttribute('data-qid');
                const idx = appState.quotes.findIndex(q => q.id === qid);
                if (idx >= 0) {
                    const removed = appState.quotes.splice(idx, 1)[0];
                    try {
                        localStorage.setItem('interfrutd_quotes', JSON.stringify(appState.quotes));
                    } catch (err) {}
                    renderQuotesModule();
                    showToast(`Cotización ${qid} (${removed.clienteNombre}) eliminada de cartera`, 'fa-trash');
                }
                return;
            }

            const btnConvert = e.target.closest('.btn-convert-quote-oc');
            if (btnConvert) {
                const qid = btnConvert.getAttribute('data-qid');
                const quote = appState.quotes.find(q => q.id === qid);
                if (!quote) return;

                quote.estado = 'CONVERTIDA_OC';
                const cli = appState.clients.find(c => c.id === quote.clienteId || c.nombre === quote.clienteNombre);
                const newOcNumber = `OC-${88440 + appState.purchaseOrders.length}`;
                const newOC = {
                    ocNumero: newOcNumber,
                    cotizacionRef: quote.id,
                    fechaRecepcion: new Date().toISOString().slice(0, 10),
                    fechaEntregaCompromiso: quote.vigenciaHasta,
                    diaSemana: (cli?.diasEntrega?.[0] || 'L'),
                    ventanaAnden: (cli?.ventanaAnden || '05:00 AM – 07:00 AM'),
                    clienteId: cli?.id || quote.clienteId || 'CLI-001',
                    clienteNombre: quote.clienteNombre,
                    segmento: quote.segmento || 'Corporativo',
                    canal: quote.listaPrecio,
                    estado: 'OC_RECIBIDA',
                    prioridad: 'ALTA',
                    bodegaAsignada: cli?.ruta || 'Ruta GAM Central (402)',
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
                    localStorage.setItem('interfrutd_purchase_orders_v2', JSON.stringify(appState.purchaseOrders));
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
    // Eventos Módulo 3: Programación Semanal, Clientes B2B & Órdenes (OC)
    // ------------------------------------------------------------------------
    document.querySelectorAll('#dayFilterGroup .chip-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#dayFilterGroup .chip-btn').forEach(b => {
                b.classList.remove('active');
                b.setAttribute('aria-pressed', 'false');
            });
            btn.classList.add('active');
            btn.setAttribute('aria-pressed', 'true');
            appState.selectedScheduleDay = btn.getAttribute('data-day') || 'ALL';
            renderOrdersAndInventoryModule();
        });
    });

    const clientSchedTbody = document.getElementById('clientScheduleTableBody');
    if (clientSchedTbody) {
        clientSchedTbody.addEventListener('click', (e) => {
            const dayBtn = e.target.closest('.day-pill-btn');
            if (dayBtn) {
                const cid = dayBtn.getAttribute('data-cli') || dayBtn.getAttribute('data-clientid');
                const daycode = dayBtn.getAttribute('data-daycode');
                const cli = appState.clients.find(c => c.id === cid);
                if (!cli || !daycode) return;

                if (!Array.isArray(cli.diasEntrega)) cli.diasEntrega = ['L', 'K', 'V'];
                const idx = cli.diasEntrega.indexOf(daycode);
                if (idx >= 0) {
                    if (cli.diasEntrega.length <= 1) {
                        showToast('Cada cliente activo debe mantener al menos 1 día de entrega asignado.', 'fa-triangle-exclamation');
                        return;
                    }
                    cli.diasEntrega.splice(idx, 1);
                } else {
                    const order = ['L', 'M', 'K', 'J', 'V', 'S'];
                    cli.diasEntrega.push(daycode);
                    cli.diasEntrega.sort((a, b) => order.indexOf(a) - order.indexOf(b));
                }

                const baseCli = B2B_CLIENTS.find(c => c.id === cid);
                if (baseCli) baseCli.diasEntrega = [...cli.diasEntrega];

                try {
                    const savedMap = {};
                    appState.clients.forEach(c => { savedMap[c.id] = c.diasEntrega; });
                    localStorage.setItem('interfrutd_client_days', JSON.stringify(savedMap));
                    localStorage.setItem('interfrutd_clients_v3', JSON.stringify(appState.clients));
                } catch (err) {}

                renderOrdersAndInventoryModule();
                showToast(`Programación de ${cli.nombre} actualizada (${cli.diasEntrega.join(', ')})`, 'fa-calendar-check');
                return;
            }

            const editBtn = e.target.closest('.btn-edit-client');
            if (editBtn) {
                const cid = editBtn.getAttribute('data-cli');
                openClientModal(cid);
                return;
            }

            const quoteBtn = e.target.closest('.btn-schedule-quote');
            if (quoteBtn) {
                const cid = quoteBtn.getAttribute('data-cli') || quoteBtn.getAttribute('data-clientid');
                const cli = appState.clients.find(c => c.id === cid);
                if (!cli) return;

                const draft = appState.activeQuoteDraft;
                draft.clienteId = cli.id;
                draft.clienteNombre = cli.nombre;
                draft.segmento = cli.segmento;
                draft.listaPrecio = cli.listaAsignada;
                draft.items.forEach(it => {
                    const cat = appState.catalog.find(x => x.sku === it.sku);
                    if (cat) {
                        it.precioLista = getBasePriceByList(cat, draft.listaPrecio);
                        it.precioFinal = Math.round(it.precioLista * (1 - (it.descuentoPct || 0) / 100));
                    }
                });

                renderQuotesModule();
                document.querySelector('.module-tab-btn[data-tab="cotizaciones"]')?.click();
                showToast(`Cliente ${cli.nombre} cargado en el Cotizador B2B`, 'fa-file-invoice-dollar');
            }
        });
    }

    // Modal de Creación / Edición de Cliente
    const btnOpenNewClient = document.getElementById('btnOpenNewClientModal');
    if (btnOpenNewClient) btnOpenNewClient.addEventListener('click', () => openClientModal(null));

    const btnQuickNewQuote = document.getElementById('btnQuickNewClientFromQuote');
    if (btnQuickNewQuote) btnQuickNewQuote.addEventListener('click', () => openClientModal(null));

    const modalClient = document.getElementById('clientFormModal');
    const btnCloseClient = document.getElementById('btnCloseClientModal');
    const btnCancelClient = document.getElementById('btnCancelClientModal');
    if (btnCloseClient && modalClient) {
        btnCloseClient.addEventListener('click', () => modalClient.classList.remove('open'));
    }
    if (btnCancelClient && modalClient) {
        btnCancelClient.addEventListener('click', () => modalClient.classList.remove('open'));
    }
    if (modalClient) {
        modalClient.addEventListener('click', (e) => {
            if (e.target === modalClient) modalClient.classList.remove('open');
        });
    }

    const dayPillsForm = document.querySelectorAll('#clientFormDayPills .day-pill-btn');
    dayPillsForm.forEach(btn => {
        btn.addEventListener('click', () => {
            const isAct = btn.classList.contains('active');
            const activeCount = document.querySelectorAll('#clientFormDayPills .day-pill-btn.active').length;
            if (isAct && activeCount <= 1) {
                showToast('El cliente debe tener al menos 1 día de despacho seleccionado.', 'fa-triangle-exclamation');
                return;
            }
            btn.classList.toggle('active');
            btn.setAttribute('aria-pressed', btn.classList.contains('active') ? 'true' : 'false');
        });
    });

    const btnSaveClient = document.getElementById('btnSaveClientModal');
    if (btnSaveClient) {
        btnSaveClient.addEventListener('click', saveClientFormData);
    }

    document.querySelectorAll('#ocStatusFilterGroup .chip-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#ocStatusFilterGroup .chip-btn').forEach(b => {
                b.classList.remove('active');
                b.setAttribute('aria-pressed', 'false');
            });
            btn.classList.add('active');
            btn.setAttribute('aria-pressed', 'true');
            appState.ocFilterStatus = btn.getAttribute('data-ocfilter');
            renderOrdersAndInventoryModule();
        });
    });

    const ocTbody = document.getElementById('purchaseOrdersTableBody');
    if (ocTbody) {
        ocTbody.addEventListener('click', (e) => {
            const btnView = e.target.closest('.btn-view-oc');
            if (btnView) {
                const ocNum = btnView.getAttribute('data-oc');
                const oc = appState.purchaseOrders.find(x => x.ocNumero === ocNum);
                if (oc) openOcDetailModal(oc);
                return;
            }

            const btnMarkDelivered = e.target.closest('.btn-mark-delivered-oc');
            if (btnMarkDelivered) {
                const ocNum = btnMarkDelivered.getAttribute('data-oc');
                const oc = appState.purchaseOrders.find(x => x.ocNumero === ocNum);
                if (!oc) return;
                oc.estado = 'ENTREGADA';
                try {
                    localStorage.setItem('interfrutd_purchase_orders_v2', JSON.stringify(appState.purchaseOrders));
                } catch (err) {}
                renderOrdersAndInventoryModule();
                showToast(`Orden ${oc.ocNumero} marcada como RECIBIDA / ENTREGADA en andén`, 'fa-circle-check');
                return;
            }

            const btnDeleteOc = e.target.closest('.btn-delete-oc');
            if (btnDeleteOc) {
                const ocNum = btnDeleteOc.getAttribute('data-oc');
                const idx = appState.purchaseOrders.findIndex(x => x.ocNumero === ocNum);
                if (idx >= 0) {
                    const removed = appState.purchaseOrders.splice(idx, 1)[0];
                    try {
                        localStorage.setItem('interfrutd_purchase_orders_v2', JSON.stringify(appState.purchaseOrders));
                    } catch (err) {}
                    renderOrdersAndInventoryModule();
                    showToast(`Orden ${ocNum} (${removed.clienteNombre}) eliminada del pipeline`, 'fa-trash');
                }
                return;
            }

            const btnAdv = e.target.closest('.btn-advance-oc');
            if (btnAdv) {
                const ocNum = btnAdv.getAttribute('data-oc');
                const oc = appState.purchaseOrders.find(x => x.ocNumero === ocNum);
                if (!oc) return;
                const flow = ['OC_RECIBIDA', 'EN_ALISTAMIENTO', 'EN_RUTA', 'ENTREGADA'];
                const currIdx = flow.indexOf(oc.estado);
                oc.estado = currIdx >= 0 && currIdx < flow.length - 1 ? flow[currIdx + 1] : 'ENTREGADA';
                try {
                    localStorage.setItem('interfrutd_purchase_orders_v2', JSON.stringify(appState.purchaseOrders));
                } catch (err) {}
                renderOrdersAndInventoryModule();
                showToast(`Orden ${oc.ocNumero} actualizada a estado: ${oc.estado}`, 'fa-truck-fast');
            }
        });
    }

    const modalOc = document.getElementById('ocDetailModal');
    const btnCloseOc = document.getElementById('btnCloseOcModal');
    if (btnCloseOc && modalOc) {
        btnCloseOc.addEventListener('click', () => modalOc.classList.remove('open'));
        modalOc.addEventListener('click', (e) => {
            if (e.target === modalOc) modalOc.classList.remove('open');
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
                    name: '2_Programacion_Clientes',
                    data: appState.clients.map(c => ({
                        ID: c.id,
                        Cliente: c.nombre,
                        Segmento: c.segmento,
                        Lista_Asignada: c.listaAsignada,
                        Ruta: c.ruta,
                        Dias_Entrega: (c.diasEntrega || []).join('-'),
                        Ventana_Anden: c.ventanaAnden,
                        Hora_Corte: c.horaCorte,
                        Credito_Dias: c.diasCredito,
                        SLA_Meta_Pct: c.slaMeta
                    }))
                },
                {
                    name: '3_Ordenes_Compra_OC',
                    data: appState.purchaseOrders.map(o => ({
                        OC: o.ocNumero,
                        Cotizacion_Ref: o.cotizacionRef,
                        Cliente: o.clienteNombre,
                        Dia: o.diaSemana,
                        Ventana_Anden: o.ventanaAnden,
                        Ruta: o.bodegaAsignada,
                        Entrega: o.fechaEntregaCompromiso,
                        Estado: o.estado
                    }))
                },
                {
                    name: '4_Benchmark_CENADA',
                    data: appState.cenadaList
                }
            ], 'Reporte_Corporativo_Interfrutd_Pricing_Comercial.xlsx');
        });
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initStateData();
    setupEvents();
    renderActiveModule();
});
