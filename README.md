# RSExcel

![License](https://img.shields.io/npm/l/rs-excel)
![Version](https://img.shields.io/npm/v/rs-excel)
![Size](https://img.shields.io/bundlephobia/minzip/rs-excel)

**RSExcel** es una librería ligera para exportar datos a **Excel (.xlsx)** desde el navegador, con soporte para:

- ✅ **Filtros automáticos**
- ✅ **Autoajuste de columnas**
- ✅ **Múltiples hojas**
- ✅ **Sin dependencias externas en runtime** (todo en un solo archivo)
- ✅ **100% compatible con la versión gratuita de SheetJS**

Ideal para aplicaciones web que requieren reportes profesionales con formato avanzado.

---

## 🚀 Instalación

### Opción 1: Descarga directa (recomendado para frontend puro)

1. [Descarga `rs.excel.min.js`](dist/rs.excel.min.js) (después de construir el proyecto)
2. Inclúyelo en tu HTML:

```html
<script src="rs.excel.min.js"></script>
```
## Opción 2: Como módulo (con bundler)
```code
npm install rs-excel
o
import RSExcel from 'rs.excel.min';
```
## 🧪 Ejemplo de uso
▶️ Una sola hoja (simple)
```javascript
RSExcel.exportToExcel(
    [['Nombre', 'Edad'], ['Ana', 28]],
    'usuarios.xlsx',
    'Hoja1'
);
```
▶️ Múltiples hojas (avanzado)
```javascript
const libro = new RSExcel.Workbook();

libro
  .addSheet('Usuarios', [['Nombre', 'Edad'], ['Ana', 28]], {
    headerStyle: { font: { bold: true, color: '#FFFFFF' }, fill: '#2C3E50' }
  })
  .addSheet('Productos', [['Producto', 'Precio'], ['Laptop', 1200]], {
    headerStyle: { font: { bold: true }, fill: '#3498DB' }
  })
  .download('reporte.xlsx');
```
## 🎨 Auto ajustable y filtros
```text
RSExcel.exportToExcel(data, filename, sheetName, {
   autoFit: true,
   enableFilters: true
});
```
## ⚠️ Notas importantes
```text
Basado en SheetJS xlsx@0.18.5 (versión gratuita, Apache 2.0).
No requiere internet en runtime.
```
## ⚖️ Licencia
```text
RSExcel: MIT License © Rogelio Urieta Camacho (RojeruSan)
SheetJS (xlsx): incluido bajo Apache License 2.0
```
## 🙌 Autor
```text
Rogelio Urieta Camacho (RojeruSan)
Hecho con ❤️ para desarrolladores que aman el control total sobre sus exportaciones.
```