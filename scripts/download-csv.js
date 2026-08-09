// Descarga el archivo anual "Contratos de la Plataforma Integral Compras MX/CompraNet [año]"
// desde Datos Abiertos de Compras MX, usando automatización de navegador
// (el archivo se sirve vía un formulario firmado con reCAPTCHA v3, no una URL pública).
//
// Uso: node scripts/download-csv.js [año]   (por defecto: año actual)

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const DATOS_ABIERTOS_URL = 'https://comprasmx.buengobierno.gob.mx/datos-abiertos';
const OUT_DIR = path.join(__dirname, '..', 'data', 'raw');
const DOWNLOAD_TIMEOUT = 60000;

async function main() {
  const year = process.argv[2] || String(new Date().getFullYear());
  const linkTextCandidates = [
    `Contratos de la Plataforma Integral Compras MX ${year}`,
    `Contratos de la Plataforma Integral CompraNet ${year}`,
  ];

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log(`Navegando a Datos Abiertos...`);
  await page.goto(DATOS_ABIERTOS_URL, { waitUntil: 'networkidle', timeout: DOWNLOAD_TIMEOUT });
  await page.waitForTimeout(1000);

  let link = null;
  for (const text of linkTextCandidates) {
    const candidate = page.getByText(text, { exact: true }).first();
    if (await candidate.count() > 0) { link = candidate; console.log(`Encontrado: "${text}"`); break; }
  }
  if (!link) {
    await browser.close();
    throw new Error(`No se encontró el link de descarga para el año ${year}. Candidatos probados: ${linkTextCandidates.join(' | ')}`);
  }

  const downloadPromise = context.waitForEvent('download', { timeout: DOWNLOAD_TIMEOUT });
  await link.click();
  console.log('Esperando descarga...');
  const download = await downloadPromise;

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const outPath = path.join(OUT_DIR, `contratos_${year}.csv`);
  await download.saveAs(outPath);

  await browser.close();

  const stats = fs.statSync(outPath);
  console.log(`Guardado: ${outPath} (${(stats.size / 1024 / 1024).toFixed(1)} MB)`);
}

main().catch(err => {
  console.error('ERROR:', err);
  process.exit(1);
});
