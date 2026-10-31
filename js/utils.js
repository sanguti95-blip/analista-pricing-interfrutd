/**
 * js/utils.js
 * Utilidades de formato numérico, moneda (CRC ₡), notificaciones y motor de exportación Excel Avanzado (SheetJS).
 */

export function formatCRC(val, decimals = 0) {
    const n = Number(val);
    if (isNaN(n)) return '₡0';
    return '₡' + n.toLocaleString('es-CR', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}

export function formatNum(val, decimals = 0) {
    const n = Number(val);
    if (isNaN(n)) return '0';
    return n.toLocaleString('es-CR', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}

export function formatPct(val, decimals = 1) {
    const n = Number(val);
    if (isNaN(n)) return '0.0%';
    return n.toFixed(decimals) + '%';
}

export function escapeHtml(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

export function showToast(message, icon = 'fa-circle-check') {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const div = document.createElement('div');
    div.className = 'toast-msg';
    div.innerHTML = `<i class="fa-solid ${icon}" style="color: var(--brand-lime); font-size: 15px;"></i><span>${escapeHtml(message)}</span>`;
    container.appendChild(div);
    setTimeout(() => {
        div.style.opacity = '0';
        div.style.transition = 'opacity 0.25s';
        setTimeout(() => div.remove(), 260);
    }, 3200);
}

export function exportRowsToExcel(rows, sheetName, fileName) {
    if (typeof XLSX === 'undefined') {
        showToast('Cargando motor Excel SheetJS, intente en un segundo...', 'fa-spinner');
        return;
    }
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
    XLSX.writeFile(wb, fileName);
    showToast(`Archivo Excel generado: ${fileName}`, 'fa-file-excel');
}

export function exportMultiSheetExcel(sheetsConfig, fileName) {
    if (typeof XLSX === 'undefined') {
        showToast('Motor SheetJS no disponible aún.', 'fa-triangle-exclamation');
        return;
    }
    const wb = XLSX.utils.book_new();
    sheetsConfig.forEach(s => {
        const ws = XLSX.utils.json_to_sheet(s.data);
        XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0, 31));
    });
    XLSX.writeFile(wb, fileName);
    showToast(`Libro Corporativo exportado: ${fileName}`, 'fa-file-excel');
}
