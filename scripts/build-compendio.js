// Convierte el Compendio Nacional de Medicamentos (CSG, .xlsm) a un JSON
// liviano indexado por clave, para usarlo como tabla de referencia en el pipeline.
//
// Uso: node scripts/build-compendio.js

const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const SRC = path.join(__dirname, '..', 'data', 'raw', 'compendio_medicamentos.xlsm');
const OUT = path.join(__dirname, '..', 'data', 'compendio_medicamentos.json');

const CLAVE_RE = /^\d{3}\.\d{3}\.\d{4}\.\d{2}/;

function main() {
  const wb = XLSX.readFile(SRC);
  const sheetName = wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });

  const compendio = {};
  let skipped = 0;

  for (const row of rows) {
    const clave = String(row['Clave'] || '').trim();
    if (!CLAVE_RE.test(clave)) { skipped++; continue; }

    // "Grupo" puede traer varios grupos terapéuticos separados por doble salto de línea,
    // cada uno con el formato "Grupo N° X: \nNombre". Nos quedamos con la lista de nombres.
    const grupoRaw = String(row['Grupo'] || '');
    const grupos = grupoRaw
      .split(/\n\s*\n/)
      .map(g => g.replace(/Grupo N[°º]\s*\d+:\s*/i, '').trim())
      .filter(Boolean);

    compendio[clave] = {
      insumo: String(row['Insumo'] || '').trim(),
      grupos_terapeuticos: grupos,
      descripcion: String(row['Descripción'] || '').replace(/\s+/g, ' ').trim(),
    };
  }

  fs.writeFileSync(OUT, JSON.stringify(compendio, null, 2), 'utf8');
  console.log(`Compendio procesado: ${Object.keys(compendio).length} claves válidas, ${skipped} filas descartadas.`);
  console.log(`Guardado en: ${OUT}`);
}

main();
