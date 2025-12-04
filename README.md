# RSExcel

![License](https://img.shields.io/npm/l/rs-excel)
![Version](https://img.shields.io/npm/v/rs-excel)
![Size](https://img.shields.io/bundlephobia/minzip/rs-excel)

**RSExcel** es una librería ligera y autónoma para exportar datos a **Excel (.xlsx)** directamente desde el navegador, con soporte completo para:

- ✅ **Múltiples hojas**
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
  .addSheet('Usuarios', users)
  .addSheet('Productos', products)
  .download('reporte.xlsx');
```
## 🎨 Auto ajustable y filtros
```text
RSExcel.exportToExcel(data, filename, sheetName, {
   autoFit: true,
   enableFilters: true
});
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