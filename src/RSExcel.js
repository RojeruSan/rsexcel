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

function downloadWorkbook(wb, filename) {
    const uint8Array = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([uint8Array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
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

// ✅ Clase para múltiples hojas y control total
export class RSExcelWorkbook {
    constructor(options = {}) {
        this.workbook = XLSX.utils.book_new();
        this.defaultAutoFit = options.autoFit ?? true;
        this.defaultEnableFilters = options.enableFilters ?? true;
    }

    addSheet(name, aoa, options = {}) {
        if (!Array.isArray(aoa) || aoa.length === 0) {
            throw new Error('RSExcel: datos inválidos');
        }

        const ws = {};
        XLSX.utils.sheet_add_aoa(ws, aoa, { origin: 'A1' });

        const autoFit = options.autoFit ?? this.defaultAutoFit;
        const enableFilters = options.enableFilters ?? this.defaultEnableFilters;
        const userHeaderStyle = options.headerStyle || {};
        const userCellStyles = options.cellStyles || {};

        const rowCount = aoa.length;
        const colCount = aoa[0]?.length || 0;

        // Estilos de encabezado
        if (rowCount > 0 && Object.keys(userHeaderStyle).length > 0) {
            const headerStyle = this._buildStyle(userHeaderStyle);
            for (let c = 0; c < colCount; c++) {
                const ref = encodeCellRef(0, c);
                if (ws[ref]) ws[ref].s = headerStyle;
            }
        }

        // Estilos por celda (ej: { 'B2': { fill: '#FF0000' } })
        for (const [addr, style] of Object.entries(userCellStyles)) {
            if (ws[addr]) {
                ws[addr].s = this._buildStyle(style);
            }
        }

        // Autoajuste
        if (autoFit && colCount > 0) {
            const colWidths = [];
            for (let c = 0; c < colCount; c++) {
                const maxWidth = Math.max(...aoa.map(row => String(row[c] || '').length), 10);
                colWidths.push({ wch: Math.min(maxWidth + 2, 50) });
            }
            ws['!cols'] = colWidths;
        }

        // Filtros
        if (enableFilters && rowCount > 0) {
            ws['!autofilter'] = {
                ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } })
            };
        }

        ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowCount - 1, c: colCount - 1 } });
        XLSX.utils.book_append_sheet(this.workbook, ws, name.substring(0, 31));
        return this;
    }

    _buildStyle(style) {
        const out = {};
        if (style.font) {
            out.font = { ...style.font };
            if (style.font.color) out.font.color = toRgbColor(style.font.color);
        }
        if (style.fill) {
            const rgb = toRgbColor(style.fill);
            out.fill = { fgColor: rgb, bgColor: rgb, patternType: 'solid' };
        }
        if (style.alignment) out.alignment = { ...style.alignment };
        return out;
    }

    download(filename = 'workbook.xlsx') {
        if (!filename.toLowerCase().endsWith('.xlsx')) filename += '.xlsx';
        downloadWorkbook(this.workbook, filename);
    }
}

// ✅ Función simple para una sola hoja
export function exportToExcel(aoa, filename = 'archivo.xlsx', sheetName = 'Hoja1', options = {}) {
    const wb = new RSExcelWorkbook(options);
    wb.addSheet(sheetName, aoa, options).download(filename);
}

// ✅ API pública
const RSExcel = {
    exportToExcel,
    Workbook: RSExcelWorkbook
};

export default RSExcel;