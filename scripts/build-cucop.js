// Convierte el catálogo CUCoP+ (.xlsx) a un JSON liviano indexado por
// cve_cucop, para usarlo como tabla de referencia en el pipeline sin volver a
// parsear el xlsx en cada corrida. Mismo patrón que build-compendio.js: el
// archivo crudo (data/raw/cucop.xlsx) no se versiona (.gitignore), este JSON
// derivado sí -- así extract.js/validar-claves.js nunca necesitan el xlsx
// crudo en runtime (importa en particular para CI, que solo tiene lo que está
// commiteado).
//
// Uso: node scripts/build-cucop.js

const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { CLAVE_RE } = require('./lib/dataset');

const SRC = path.join(__dirname, '..', 'data', 'raw', 'cucop.xlsx');
const OUT = path.join(__dirname, '..', 'data', 'cucop_medicamentos.json');

// Para toda entrada de la partida 25301 (medicamentos), la columna
// "DESCRIPCIÓN" del catálogo trae la clave del Compendio CSG como prefijo --
// es estático (no depende de qué institución compre), así que sirve de
// respaldo para resolver `clave` cuando la descripción del contrato no la
// incluye directamente (ver Metodologia.md §3.4/§5.2).
//
// No todas las entradas traen ese prefijo: ~23% de las 25301 (sobre todo
// altas recientes al catálogo) solo tienen el nombre del producto, sin clave.
// Se guardan igual, con `clave: null`, para que su `descripcion` siga siendo
// utilizable (ej. para sanear un `producto` degenerado -- Metodologia.md
// §5.3.2) aunque no aporte clave.
function main() {
  const wb = XLSX.readFile(SRC);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });

  const map = {};
  let sinClave = 0;
  for (const row of rows) {
    if (String(row['PARTIDA ESPECÍFICA'] || '').trim() !== '25301') continue;
    const cveCucop = String(row['CLAVE CUCoP +'] || '').trim();
    const descripcionCruda = String(row['DESCRIPCIÓN'] || '').trim();
    if (!cveCucop || !descripcionCruda) continue;
    const m = descripcionCruda.match(CLAVE_RE);
    if (m) {
      map[cveCucop] = { clave: `${m[1]}.${m[2]}.${m[3]}.${m[4]}`, descripcion: m[5].trim() };
    } else {
      map[cveCucop] = { clave: null, descripcion: descripcionCruda };
      sinClave++;
    }
  }

  fs.writeFileSync(OUT, JSON.stringify(map, null, 2), 'utf8');
  console.log(`Catálogo CUCoP+ procesado: ${Object.keys(map).length} claves (${sinClave} sin prefijo de clave CSG en su ficha).`);
  console.log(`Guardado en: ${OUT}`);
}

main();
