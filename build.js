// build.js
const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

esbuild.build({
    entryPoints: [path.resolve(__dirname, 'src/index.js')], // ← usa index.js
    bundle: true,
    minify: true,
    legalComments: 'none',
    platform: 'browser',
    format: 'iife',
    outfile: path.resolve(__dirname, 'dist/rs.excel.min.js'),
    define: {
        'process.env.NODE_ENV': '"production"'
    }
}).then(() => {
    const banner = `/**
 * RSExcel - MIT License © ${new Date().getFullYear()} Rogelio Urieta Camacho
 * Bundles SheetJS (Apache 2.0)
 */
`;
    const file = path.resolve(__dirname, 'dist/rs.excel.min.js');
    const content = fs.readFileSync(file, 'utf8');
    fs.writeFileSync(file, banner + content);
    console.log('✅ RSExcel listo. Usa: RSExcel.exportToExcel(...)');
}).catch(err => {
    console.error('❌ Error:', err);
    process.exit(1);
});