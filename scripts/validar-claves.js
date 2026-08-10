// Sanity check post-corrida: TODA la lógica de confiabilidad de `clave` vive
// aquí (no en extract.js, que solo agarra una clave rápido por el medio que
// se pueda). Agrupa TODOS los registros de docs/data.json por producto
// canónico (nombre + dosis normalizados -- no texto exacto, para que
// pequeñas diferencias de redacción entre instituciones no separen lo que es
// el mismo producto) y, por grupo, decide una única clave "ganadora" que se
// propaga a todas las instancias -- así un mismo producto no puede terminar
// con dos claves distintas según qué contrato lo haya traído.
//
// Prioridad por grupo (se detiene en el primer paso que encuentre algo):
//   1. Algún miembro con clave extraída directo de la descripción del
//      contrato -- la fuente más confiable que existe, se usa tal cual.
//   2. Algún miembro con clave vía CUCoP que además AUTOVERIFICA: la dosis/
//      presentación que ese código declara en el propio catálogo CUCoP
//      coincide (>=2 números) con lo que el contrato describe. Se comprobó
//      que las instituciones a veces usan un cve_cucop que no corresponde a
//      lo que compraron (Metodologia.md §7/§9.1, casos Abacavir, Raltegravir,
//      Agua Inyectable) -- sin esto se arrastraría la clave equivocada.
//   3. Compendio Nacional local (data/compendio_medicamentos.json), buscando
//      por nombre -- gratis, sin red.
//   4. Si nada de lo anterior aplica: cualquier clave vía CUCoP que NO haya
//      pasado la autoverificación se VACÍA explícitamente (no se deja un
//      valor que ya se sabe incorrecto solo porque no hay reemplazo mejor).
//      Los registros que ya estaban en null se quedan en null. Todo esto
//      queda documentado en docs/data.correcciones.json para revisión manual.
//
// Búsqueda externa (vademecum.es/cnis, Comp26042025.pdf) NO está automatizada
// aquí por costo/beneficio bajo (~2.5s por consulta, ~2% de aciertos medido
// en una muestra real) -- queda documentada como alternativa manual en
// Limitaciones.md, no como parte de este pipeline.
//
// Solo se aplican correcciones de ALTA CONFIANZA (nombre coincide Y al menos
// 2 números coinciden -- típicamente dosis + tamaño de envase). Todo lo
// demás queda documentado en el reporte, nunca se sobreescribe a ciegas.
//
// Uso:
//   node scripts/validar-claves.js

const fs = require('fs');
const path = require('path');
const { guardarExcel } = require('./lib/dataset');

const DATA_PATH = path.join(__dirname, '..', 'docs', 'data.json');
const REPORTE_PATH = path.join(__dirname, '..', 'docs', 'data.correcciones.json');
const EXCEL_PATH = path.join(__dirname, '..', 'docs', 'data.xlsx');
const COMPENDIO_PATH = path.join(__dirname, '..', 'data', 'compendio_medicamentos.json');
const CUCOP_PATH = path.join(__dirname, '..', 'data', 'cucop_medicamentos.json');

function normalizarTexto(s) {
  return String(s || '')
    .toUpperCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[.,;:()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extraerNumeros(s) {
  const matches = normalizarTexto(s).match(/\d+(\.\d+)?/g) || [];
  return [...new Set(matches.map(Number))].sort((a, b) => a - b);
}

function primeraPalabraClave(s) {
  const norm = normalizarTexto(s);
  const palabras = norm.split(' ').filter(w => w.length > 3 && !/^\d+$/.test(w));
  return palabras[0] || norm.split(' ')[0] || '';
}

// Clave canónica de "producto" = nombre (primera palabra significativa) +
// huella numérica completa (todas las dosis/envases mencionados, ordenados).
function claveCanonica(producto) {
  return primeraPalabraClave(producto) + '|' + extraerNumeros(producto).join(',');
}

function contarCoincidenciasNumericas(numsA, numsB) {
  const setB = new Set(numsB);
  let n = 0;
  for (const x of numsA) if (setB.has(x)) n++;
  return n;
}

// Precomputado por scripts/build-cucop.js a partir del catálogo CUCoP+ --
// ver ahí. No todas las entradas traen clave CSG (~23% de las 25301, ver
// build-cucop.js); se guardan igual, con `clave: null`, para que
// `cucopEsValido` pueda seguir usando su `descripcion`.
function loadCucopMap() {
  return JSON.parse(fs.readFileSync(CUCOP_PATH, 'utf8'));
}

// Solo infiere para registros generados antes de que clave_fuente existiera
// como campo (corridas previas a 2026-08-08). Si el campo ya existe, se
// respeta tal cual -- incluye valores de corridas anteriores de este mismo
// script (validacion_nombre_local, propagacion_confiable, etc.), que ya
// pasaron por esta verificación y no hay que volver a derivar.
function inferirFuente(registro, cucopMap) {
  if (registro.clave_fuente !== undefined) return registro.clave_fuente;
  if (!registro.clave) return null;
  if (registro.cve_cucop && cucopMap[registro.cve_cucop] && cucopMap[registro.cve_cucop].clave === registro.clave) return 'cucop';
  return 'descripcion';
}

// ¿Esta fuente ya es evidencia confiable de por sí (sin volver a verificar)?
// 'cucop' es la única que necesita la autoverificación aparte (cucopEsValido)
// porque es la única que nunca pasó por ningún chequeo. Todo lo demás (la
// descripción propia del contrato, o el resultado de una corrida anterior de
// este mismo script) ya fue validado en algún punto.
function esFuenteConfiable(fuente) {
  return fuente !== null && fuente !== undefined && fuente !== 'cucop';
}

// La clave recuperada vía CUCoP solo se confía si la dosis/presentación que
// el propio catálogo declara para ese código coincide con lo que el contrato
// describe. Si el contrato no trae ningún número no se puede verificar, así
// que se confía por defecto (no hay evidencia de que esté mal).
function cucopEsValido(producto, cveCucop, cucopMap) {
  const entry = cucopMap[cveCucop];
  if (!entry) return false;
  const numsProducto = extraerNumeros(producto);
  if (numsProducto.length === 0) return true;
  return contarCoincidenciasNumericas(numsProducto, extraerNumeros(entry.descripcion)) >= 2;
}

function buscarEnCompendioLocal(producto, compendio) {
  const numsProducto = extraerNumeros(producto);
  const nombreProducto = normalizarTexto(producto);
  let mejorClave = null;
  let mejorScore = 0;
  for (const [clave, entry] of Object.entries(compendio)) {
    const nombreEntry = normalizarTexto(entry.insumo);
    const primeraPalabraEntry = nombreEntry.split(' ')[0];
    if (primeraPalabraEntry.length > 3 && !nombreProducto.includes(primeraPalabraEntry)) continue;
    const numsEntry = extraerNumeros(entry.descripcion);
    const score = contarCoincidenciasNumericas(numsProducto, numsEntry);
    if (score > mejorScore) { mejorScore = score; mejorClave = clave; }
  }
  return mejorScore >= 2 ? { clave: mejorClave, score: mejorScore } : null;
}

function asignar(data, i, clave, compendio, fuente, corregidos, representante) {
  if (data[i].clave === clave && data[i].clave_fuente === fuente) return;
  corregidos.push({
    producto: representante.slice(0, 100),
    codigo_contrato: data[i].codigo_contrato,
    clave_anterior: data[i].clave,
    clave_nueva: clave,
    fuente,
  });
  data[i].clave = clave;
  data[i].grupo_terapeutico = clave && compendio[clave] ? compendio[clave].grupos_terapeuticos : null;
  data[i].clave_fuente = fuente;
}

function main() {
  const data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
  const compendio = JSON.parse(fs.readFileSync(COMPENDIO_PATH, 'utf8'));
  const cucopMap = loadCucopMap();

  console.log('Agrupando TODOS los registros por producto canónico (nombre + dosis normalizados)...');
  const grupos = new Map(); // claveCanonica -> [indices en data]
  for (let i = 0; i < data.length; i++) {
    const key = claveCanonica(data[i].producto);
    if (!grupos.has(key)) grupos.set(key, []);
    grupos.get(key).push(i);
  }
  console.log(`${grupos.size} productos canónicos únicos (${data.length} registros).`);

  const corregidos = [];
  const sinResolver = [];
  const contadores = { por_descripcion: 0, por_cucop_valido: 0, por_local: 0, vaciados_cucop_invalido: 0, sin_resolver: 0 };

  for (const [, indices] of grupos.entries()) {
    const representante = indices.map(i => data[i].producto).sort((a, b) => b.length - a.length)[0];

    const miembros = indices.map(i => {
      const fuente = inferirFuente(data[i], cucopMap);
      const valido = esFuenteConfiable(fuente) ? true
        : fuente === 'cucop' ? cucopEsValido(data[i].producto, data[i].cve_cucop, cucopMap)
        : false;
      return { i, fuente, valido };
    });

    // Prioridad 1: miembro con clave de la descripción del contrato -- la
    // fuente más confiable que existe, gana siempre que esté presente.
    const conDescripcion = miembros.find(m => m.fuente === 'descripcion');
    // Prioridad 2: cualquier otra fuente ya confiable (validada en esta
    // corrida o en una anterior) o CUCoP que autoverifica en esta pasada.
    const conConfiable = miembros.find(m => m.valido);

    let claveGanadora = null, fuenteGanadora = null;
    if (conDescripcion) {
      claveGanadora = data[conDescripcion.i].clave;
      fuenteGanadora = 'descripcion';
      contadores.por_descripcion++;
    } else if (conConfiable) {
      claveGanadora = data[conConfiable.i].clave;
      fuenteGanadora = data[conConfiable.i].clave_fuente !== undefined ? data[conConfiable.i].clave_fuente : conConfiable.fuente;
      contadores.por_cucop_valido++;
    } else {
      const local = buscarEnCompendioLocal(representante, compendio);
      if (local) {
        claveGanadora = local.clave;
        fuenteGanadora = 'validacion_nombre_local';
        contadores.por_local++;
      }
    }

    if (claveGanadora) {
      for (const i of indices) asignar(data, i, claveGanadora, compendio, fuenteGanadora, corregidos, representante);
    } else {
      // Nada encontrado: vaciar explícitamente cualquier clave vía CUCoP que
      // no haya autoverificado -- no dejar un valor que ya se sabe mal.
      let huboVaciado = false;
      for (const m of miembros) {
        if (m.fuente === 'cucop' && !m.valido) {
          asignar(data, m.i, null, compendio, null, corregidos, representante);
          huboVaciado = true;
        }
      }
      if (huboVaciado) contadores.vaciados_cucop_invalido++;
      contadores.sin_resolver++;
      sinResolver.push({ producto: representante.slice(0, 100), registros_afectados: indices.length });
    }
  }

  fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2), 'utf8');
  guardarExcel(EXCEL_PATH, data);
  fs.writeFileSync(REPORTE_PATH, JSON.stringify({
    generado: new Date().toISOString(),
    resumen: { productos_canonicos: grupos.size, registros: data.length, registros_corregidos: corregidos.length, ...contadores },
    corregidos,
    sin_resolver: sinResolver,
  }, null, 2), 'utf8');

  console.log('\n--- Resumen ---');
  console.log(`Productos canónicos evaluados: ${grupos.size}`);
  console.log(`Grupos resueltos por descripción del contrato: ${contadores.por_descripcion}`);
  console.log(`Grupos resueltos por CUCoP autoverificado: ${contadores.por_cucop_valido}`);
  console.log(`Grupos resueltos por compendio local: ${contadores.por_local}`);
  console.log(`Grupos sin resolver: ${contadores.sin_resolver} (de esos, ${contadores.vaciados_cucop_invalido} tenían clave vía CUCoP inválida que se vació)`);
  console.log(`Registros corregidos/actualizados en total: ${corregidos.length}`);
  console.log(`Reporte: ${REPORTE_PATH}`);
}

main();
