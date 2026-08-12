// Auditoría rápida de docs/data.json + docs/data.errores.json, para
// verificar el estado del dataset sin tener que confiar ciegamente en el
// resumen que imprime una corrida de extract.js. Pensado para correrse a
// mano en cualquier momento, y para que el agente lo corra antes/después de
// cualquier operación que filtre o borre registros.
//
// Uso: node scripts/audit-dataset.js [--out docs/data.json]
//
// Nace de la sesión 2026-08-10: un "reset" de 2025 mal hecho (filtrando por
// prefijo de codigo_contrato, que NO indica el CSV de origen) borró 1,525
// registros legítimos de 2026 sin que nada lo detectara hasta que el usuario
// preguntó. Todo lo que hay abajo son chequeos que habrían mostrado ese
// problema de inmediato.

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function parseArgs(argv) {
  const rest = argv.slice(2);
  const outIdx = rest.indexOf('--out');
  return { out: outIdx !== -1 ? rest[outIdx + 1] : path.join(__dirname, '..', 'docs', 'data.json') };
}

function contarPorOrigen(registros) {
  const conteo = {};
  for (const r of registros) {
    const o = r.origen || '(sin origen)';
    conteo[o] = (conteo[o] || 0) + 1;
  }
  return conteo;
}

function main() {
  const args = parseArgs(process.argv);
  const errPath = args.out.replace(/\.json$/, '.errores.json');

  const data = JSON.parse(fs.readFileSync(args.out, 'utf8'));
  const errores = fs.existsSync(errPath) ? JSON.parse(fs.readFileSync(errPath, 'utf8')) : [];

  console.log(`=== Auditoría de ${args.out} ===\n`);

  // 1. Conteo por origen -- lo primero que hay que mirar tras cualquier
  // reset/filtro/merge.
  const porOrigen = contarPorOrigen(data);
  console.log(`Total: ${data.length} registros`);
  for (const [origen, n] of Object.entries(porOrigen).sort()) {
    console.log(`  origen "${origen}": ${n}`);
  }

  // 2. Registros sin origen -- dato legado o un bug de extracción que no lo
  // está poblando.
  const sinOrigen = data.filter(r => !r.origen).length;
  if (sinOrigen > 0) {
    console.log(`\n⚠ ${sinOrigen} registros SIN origen (dato legado o bug de extracción).`);
  }

  // 3. Comparación contra el último commit -- la que habría atrapado el
  // incidente de esta sesión al instante: cualquier origen que BAJE de
  // conteo respecto a HEAD es sospechoso (salvo que se haya pedido
  // explícitamente reiniciar ese origen).
  try {
    const headRaw = execSync(`git show HEAD:${path.relative(path.join(__dirname, '..'), args.out).replace(/\\/g, '/')}`, { cwd: path.join(__dirname, '..'), maxBuffer: 1024 * 1024 * 200 }).toString('utf8');
    const headData = JSON.parse(headRaw);
    const headPorOrigen = contarPorOrigen(headData);
    console.log(`\n=== Comparación contra HEAD (último commit) ===`);
    console.log(`HEAD: ${headData.length} registros`);
    const todosOrigenes = new Set([...Object.keys(porOrigen), ...Object.keys(headPorOrigen)]);
    for (const origen of [...todosOrigenes].sort()) {
      const antes = headPorOrigen[origen] || 0;
      const ahora = porOrigen[origen] || 0;
      const delta = ahora - antes;
      const marca = delta < 0 ? ' ⚠ BAJÓ' : delta > 0 ? ' (creció)' : '';
      console.log(`  origen "${origen}": HEAD=${antes} -> ahora=${ahora} (${delta >= 0 ? '+' : ''}${delta})${marca}`);
    }
  } catch (e) {
    console.log(`\n(No se pudo comparar contra HEAD: ${e.message.split('\n')[0]})`);
  }

  // 4. Duplicados exactos por codigo_contrato + cve_cucop -- informativo, NO
  // necesariamente un error (hay casos catálogo legítimos ya documentados en
  // Limitaciones.md donde esto pasa a propósito).
  const claves = new Map();
  for (const r of data) {
    const k = `${r.codigo_contrato}|${r.cve_cucop}`;
    claves.set(k, (claves.get(k) || 0) + 1);
  }
  const duplicados = [...claves.values()].filter(n => n > 1).length;
  console.log(`\nCombinaciones codigo_contrato+cve_cucop repetidas: ${duplicados} (ver Limitaciones.md antes de asumir que es un bug).`);

  // 5. Errores acumulados -- por origen, y cuántos ya están en fallo
  // permanente real (MAX_INTENTOS en extract.js, hoy 3) por contrato.
  console.log(`\n=== ${errPath} ===`);
  console.log(`Total: ${errores.length} errores`);
  const erroresPorOrigen = contarPorOrigen(errores);
  for (const [origen, n] of Object.entries(erroresPorOrigen).sort()) {
    console.log(`  origen "${origen}": ${n}`);
  }
  const intentosPorContrato = new Map();
  for (const e of errores) {
    const key = `${e.hash}|${e.codigoContrato}`;
    intentosPorContrato.set(key, (intentosPorContrato.get(key) || 0) + 1);
  }
  const conTresOMas = [...intentosPorContrato.values()].filter(n => n >= 3).length;
  console.log(`Contratos con >= 3 intentos fallidos (fallo permanente, ya no se reintentan solos): ${conTresOMas}`);

  console.log('\n=== Fin auditoría ===');
}

main();
