// src/RSExcel.js
// RSExcel - Advanced Excel export with full styling support
// MIT License © Rogelio Urieta Camacho (RojeruSan)

import * as XLSX from 'xlsx';

function toRgbColor(color) {
    if (!color) return undefined;
    if (typeof color === 'object' && color.rgb) return color;
    let hex = color.replace(/^#/, '');
    if (hex.length === 3) {
        hex = hex.split('').map(c => c + c).join('');
    }
    return { rgb: hex.toUpperCase() };
}

function encodeCellRef(r, c) {
    return XLSX.utils.encode_cell({ r, c });
}

// ✅ Descarga manual que respeta estilos
function downloadWorkbook(wb, filename) {
    // Usar write con bookType: 'xlsx' y type: 'array'
    const uint8Array = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([uint8Array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    // Descargar
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }, 0);
}

function exportToExcel(aoa, filename = 'archivo.xlsx', sheetName = 'Hoja1', options = {}) {
    if (!Array.isArray(aoa) || aoa.length === 0) {
        throw new Error('RSExcel: datos inválidos');
    }

    const autoFit = options.autoFit ?? true;
    const enableFilters = options.enableFilters ?? true;
    const userHeaderStyle = options.headerStyle || {};

    const ws = {};
    XLSX.utils.sheet_add_aoa(ws, aoa, { origin: 'A1' });

    const rowCount = aoa.length;
    const colCount = aoa[0]?.length || 0;

    // Aplicar estilos al encabezado
    if (rowCount > 0 && Object.keys(userHeaderStyle).length > 0) {
        const headerStyle = {};

        if (userHeaderStyle.font) {
            headerStyle.font = { ...userHeaderStyle.font };
            if (userHeaderStyle.font.color) {
                headerStyle.font.color = toRgbColor(userHeaderStyle.font.color);
            }
        }

        if (userHeaderStyle.fill) {
            const rgb = toRgbColor(userHeaderStyle.fill);
            headerStyle.fill = {
                fgColor: rgb,
                bgColor: rgb,
                patternType: 'solid'
            };
        }

        for (let c = 0; c < colCount; c++) {
            const ref = encodeCellRef(0, c);
            if (ws[ref]) {
                ws[ref].s = headerStyle;
            }
        }
    }

    // Autoajuste
    if (autoFit && colCount > 0) {
        const colWidths = [];
        for (let c = 0; c < colCount; c++) {
            const maxWidth = Math.max(
                ...aoa.map(row => String(row[c] || '').length),
                10
            );
            colWidths.push({ wch: Math.min(maxWidth + 2, 50) });
        }
        ws['!cols'] = colWidths;
    }

    // Filtros
    if (enableFilters && rowCount > 0) {
        ws['!autofilter'] = {
            ref: XLSX.utils.encode_range({
                s: { r: 0, c: 0 },
                e: { r: 0, c: colCount - 1 }
            })
        };
    }

    ws['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: rowCount - 1, c: colCount - 1 }
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));

    if (!filename.toLowerCase().endsWith('.xlsx')) {
        filename += '.xlsx';
    }

    // ✅ Usar descarga manual (respeta estilos)
    downloadWorkbook(wb, filename);
}

const RSExcel = { exportToExcel };
export default RSExcel;