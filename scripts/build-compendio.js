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

// Catálogo oficial de grupos terapéuticos del CSG, indexado por número de
// grupo (estable) en vez de por el texto libre que lo acompaña en la celda
// "Grupo" del archivo fuente. Se usa el número como clave de verdad porque el
// texto libre trae erratas/variantes inconsistentes entre renglones del mismo
// archivo -- verificado contra el compendio completo: cada número (1-23)
// tiene un único nombre real detrás de sus erratas, ej. el grupo 23 aparece
// como "Cuidados Paliativos" / "Cuidados paliativos" / "Cuidados palIativos"
// y el grupo 20 como "Reumatología y Traumatología" / "Reumatologia y Trauma"
// / "Reumatologia y Traumatologia" según el renglón.
const GRUPOS_TERAPEUTICOS = {
  '1': 'Analgesia',
  '2': 'Anestesia',
  '3': 'Cardiología',
  '4': 'Dermatología',
  '5': 'Endocrinología y Metabolismo',
  '6': 'Enfermedades Infecciosas y Parasitarias',
  '7': 'Enfermedades Inmunoalérgicas',
  '8': 'Gastroenterología',
  '9': 'Gineco-obstetricia',
  '10': 'Hematología',
  '11': 'Intoxicaciones',
  '12': 'Nefrología y Urología',
  '13': 'Neumología',
  '14': 'Neurología',
  '15': 'Oftalmología',
  '16': 'Oncología',
  '17': 'Otorrinolaringología',
  '18': 'Planificación Familiar',
  '19': 'Psiquiatría',
  '20': 'Reumatología y Traumatología',
  '21': 'Soluciones Electrolíticas y Sustitutos del Plasma',
  '22': 'Vacunas, Toxoides, Inmunoglobulinas, Antitoxinas',
  '23': 'Cuidados Paliativos',
};

// Marca el inicio de cada grupo dentro de la celda "Grupo": "Grupo N° X:",
// "Grupo Nº X.", "Grupo Nº X°:", "Cuidados N° X:" (typo visto en el archivo
// real por "Grupo"), con o sin espacio antes del número. Los separadores
// entre grupos NO son consistentes en el archivo (a veces doble salto de
// línea, a veces uno solo, a veces ninguno) -- por eso se separa por este
// marcador en vez de por saltos de línea, que fallaba en filas donde dos
// grupos quedaban pegados en un solo string sin split correcto.
const MARKER_RE = /(?:Grupo|Cuidados)\s*N[°ºo]?\.?\s*(\d+)\s*[°º]?\s*[:.]\s*/gi;

function parseGrupos(grupoRaw) {
  if (!grupoRaw) return [];
  const marcadores = [...grupoRaw.matchAll(MARKER_RE)];
  const grupos = [];
  for (let i = 0; i < marcadores.length; i++) {
    const numero = marcadores[i][1];
    const inicioTexto = marcadores[i].index + marcadores[i][0].length;
    const finTexto = i + 1 < marcadores.length ? marcadores[i + 1].index : grupoRaw.length;
    const textoLibre = grupoRaw.slice(inicioTexto, finTexto).replace(/\s+/g, ' ').trim();
    const canonico = GRUPOS_TERAPEUTICOS[numero];
    if (!canonico) {
      console.warn(`Grupo terapéutico Nº ${numero} no está en GRUPOS_TERAPEUTICOS (texto de origen: "${textoLibre}") -- revisar y agregarlo en scripts/build-compendio.js.`);
    }
    grupos.push(canonico || textoLibre);
  }
  return [...new Set(grupos)].filter(Boolean);
}

function main() {
  const wb = XLSX.readFile(SRC);
  const sheetName = wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });

  const compendio = {};
  let skipped = 0;

  for (const row of rows) {
    const clave = String(row['Clave'] || '').trim();
    if (!CLAVE_RE.test(clave)) { skipped++; continue; }

    compendio[clave] = {
      insumo: String(row['Insumo'] || '').trim(),
      grupos_terapeuticos: parseGrupos(String(row['Grupo'] || '')),
      descripcion: String(row['Descripción'] || '').replace(/\s+/g, ' ').trim(),
    };
  }

  fs.writeFileSync(OUT, JSON.stringify(compendio, null, 2), 'utf8');
  console.log(`Compendio procesado: ${Object.keys(compendio).length} claves válidas, ${skipped} filas descartadas.`);
  console.log(`Guardado en: ${OUT}`);
}

main();
