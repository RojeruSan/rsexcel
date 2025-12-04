# RSExcel

![License](https://img.shields.io/npm/l/rs-excel)
![Version](https://img.shields.io/npm/v/rs-excel)
![Size](https://img.shields.io/bundlephobia/minzip/rs-excel)

**RSExcel** es una librería ligera y autónoma para exportar datos a **Excel (.xlsx)** directamente desde el navegador, con soporte completo para:

- ✅ **Múltiples hojas**
- ✅ **Estilos personalizados** (color de fondo, color de texto, negrita, alineación)
- ✅ **Filtros automáticos**
- ✅ **Autoajuste de columnas**
- ✅ **Sin dependencias externas en runtime** (todo incluido en un solo archivo)
- ✅ **Licencia MIT**

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
```javascript
const users = [
  ['Nombre', 'Edad', 'Ciudad'],
  ['Ana', 28, 'Madrid'],
  ['Luis', 34, 'Barcelona']
];

const products = [
  ['Producto', 'Precio'],
  ['Laptop', 1200],
  ['Mouse', 25]
];

const excel = new RSExcel({
  autoFit: true,
  enableFilters: true
});

excel
  .addSheet('Usuarios', users, {
    headerStyle: {
      font: { bold: true, color: '#FFFFFF' },
      fill: '#2C3E50'
    },
    styles: {
      'B2': { font: { color: '#E74C3C', bold: true } }
    }
  })
  .addSheet('Productos', products, {
    headerStyle: {
      font: { bold: true, color: '#27AE60' },
      fill: '#F8F9FA'
    }
  })
  .download('reporte.xlsx');
```
## 🛠️ API
```text
new RSExcel(options)

options.autoFit (boolean, default: true)
options.enableFilters (boolean, default: true)

.addSheet(name, data, options)

name: nombre de la hoja (máx. 31 caracteres)
data: array 2D ([['A1','B1'], ['A2','B2']])
options:
autoFit: sobrescribe la configuración global
enableFilters: sobrescribe la configuración global
headerStyle: estilo aplicado a la fila 0
styles: objeto con estilos por celda ('A1', 'B2', etc.) o por rangos

.download(filename)

Genera y descarga el archivo Excel.
```
## 🎨 Estilos soportados
```text
{
  fill: '#FF5733', // color de fondo (hex)
  font: {
    bold: true,
    italic: false,
    color: '#FFFFFF',
    name: 'Arial',
    sz: 12
  },
  alignment: {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
  }
}
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