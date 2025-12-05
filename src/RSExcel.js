// src/RSExcel.js
// RSExcel - Advanced Excel export with full styling support
// MIT License © Rogelio Urieta Camacho (RojeruSan)
// src/RSExcel.js
import * as XLSX from 'xlsx';

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

// ✅ Extraer datos de una tabla HTML
function tableToAoa(table) {
    if (typeof table === 'string') {
        table = document.querySelector(table);
    }
    if (!table || table.tagName !== 'TABLE') {
        throw new Error('RSExcel: se requiere un elemento <table> válido');
    }

    const aoa = [];
    const rows = table.querySelectorAll('tr');

    rows.forEach(row => {
        const cells = row.querySelectorAll('th, td');
        const rowData = [];
        cells.forEach(cell => {
            // Obtener texto (ignorar HTML interno)
            let text = cell.textContent || cell.innerText || '';
            // Eliminar espacios redundantes
            text = text.replace(/\s+/g, ' ').trim();
            rowData.push(text);
        });
        aoa.push(rowData);
    });

    return aoa;
}

// ✅ Clase principal
export class RSExcelWorkbook {
    constructor(options = {}) {
        this.workbook = XLSX.utils.book_new();
        this.defaultAutoFit = options.autoFit ?? true;
        this.defaultEnableFilters = options.enableFilters ?? true;
    }

    // ✅ Añadir hoja desde datos estructurados
    addSheet(name, data, columns = [], options = {}) {
        if (!Array.isArray(data)) {
            throw new Error('RSExcel: data must be an array');
        }

        const autoFit = options.autoFit ?? this.defaultAutoFit;
        const enableFilters = options.enableFilters ?? this.defaultEnableFilters;
        const exportWithRowNumbers = options.exportWithRowNumbers ?? false;
        const rowNumberText = options.rowNumberText || 'N°';
        const summary = options.summary;

        const headers = columns.map(col => col.title || col.name || col.key);
        if (exportWithRowNumbers) headers.unshift(rowNumberText);

        const rows = data.map((row, idx) => {
            let rowData = columns.map(col => String(row[col.key] ?? ''));
            if (exportWithRowNumbers) rowData.unshift(idx + 1);
            return rowData;
        });

        if (summary && Array.isArray(summary.values)) {
            const summaryRow = [summary.label || 'Total', ...summary.values];
            if (!exportWithRowNumbers) summaryRow.shift(); // quitar label si no hay N°
            rows.push(summaryRow);
        }

        this._addSheetFromAoa(name, [headers, ...rows], { autoFit, enableFilters });
        return this;
    }

    // ✅ Añadir hoja desde una tabla HTML (cadena de selector o elemento)
    addSheetFromTable(name, tableSelector, options = {}) {
        const aoa = tableToAoa(tableSelector);
        const autoFit = options.autoFit ?? this.defaultAutoFit;
        const enableFilters = options.enableFilters ?? this.defaultEnableFilters;
        this._addSheetFromAoa(name, aoa, { autoFit, enableFilters });
        return this;
    }

    // ✅ Lógica interna para crear hoja desde aoa
    _addSheetFromAoa(name, aoa, options) {
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const colCount = aoa[0]?.length || 0;
        const rowCount = aoa.length;

        // Autoajuste
        if (options.autoFit && colCount > 0) {
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
        if (options.enableFilters && rowCount > 0) {
            ws['!autofilter'] = {
                ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } })
            };
        }

        ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowCount - 1, c: colCount - 1 } });
        XLSX.utils.book_append_sheet(this.workbook, ws, name.substring(0, 31));
    }

    download(filename = 'workbook.xlsx') {
        if (!filename.toLowerCase().endsWith('.xlsx')) filename += '.xlsx';
        downloadWorkbook(this.workbook, filename);
    }
}

// ✅ Atajos
export function exportToExcel(data, columns, filename, sheetName, options) {
    const wb = new RSExcelWorkbook(options);
    wb.addSheet(sheetName, data, columns, options).download(filename);
}

export function exportTableToExcel(tableSelector, filename = 'tabla.xlsx', sheetName = 'Hoja1', options = {}) {
    const wb = new RSExcelWorkbook(options);
    wb.addSheetFromTable(sheetName, tableSelector, options).download(filename);
}

// ✅ API pública
const RSExcel = {
    exportToExcel,
    exportTableToExcel,
    Workbook: RSExcelWorkbook
};

export default RSExcel;