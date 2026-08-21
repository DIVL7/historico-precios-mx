// Sanity check post-corrida: TODA la lógica de confiabilidad de `clave` vive
// aquí (no en extract.js, que solo agarra una clave rápido por el medio que
// se pueda). Evalúa CADA registro de docs/data.json de forma AISLADA -- nunca
// le presta evidencia de otro registro, ni siquiera de uno con el mismo
// nombre+dosis. Decisión tomada el 2026-08-17 (ver Metodologia.md): la
// versión anterior agrupaba por "producto canónico" y propagaba una clave
// ganadora a todos los miembros del grupo, pero eso no garantiza que la
// clave prestada sea correcta para CADA instancia -- se prefiere dejar un
// registro sin resolver antes que asignarle un dato que no está verificado
// para su propio contrato.
//
// Prioridad por registro (se detiene en el primer paso que encuentre algo):
//   1. Ya tiene una fuente confiable asentada (`clave_fuente` distinto de
//      `null`/`'cucop'` -- típicamente `descripcion`, extraída directo del
//      propio contrato, o el resultado de una corrida anterior de este mismo
//      script) -- no se toca, ya fue validado en algún punto.
//   2. Clave vía CUCoP que además AUTOVERIFICA: la dosis/presentación que ese
//      código declara en el propio catálogo CUCoP coincide (>=2 números) con
//      lo que EL PROPIO contrato describe -- o, si cualquiera de los dos
//      lados no trae ningún número que comparar, se confía por defecto (no
//      hay evidencia de que esté mal, ver cucopEsValido). Se comprobó que las
//      instituciones a veces usan un cve_cucop que no corresponde a lo que
//      compraron -- sin este chequeo, se arrastraría la clave equivocada.
//   3. Compendio Nacional local (data/compendio_medicamentos.json), buscando
//      con el texto del PROPIO registro -- gratis, sin red.
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
// 2 números coinciden -- típicamente dosis + tamaño de envase). El umbral se
// evaluó bajar a 1 ahora que cada registro se evalúa con su propio texto (en
// vez del texto más largo de un grupo entero) y se descartó a propósito:
// prioriza que el dato sea correcto antes que estar completo. Todo lo que no
// alcanza el umbral queda documentado en el reporte, nunca se sobreescribe a
// ciegas.
//
// Uso:
//   node scripts/validar-claves.js

const fs = require('fs');
const path = require('path');
const { guardarExcel, quitarPrefijoClave } = require('./lib/dataset');

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

// `producto` se guarda 100% crudo desde extract.js (decisión 2026-08-21, ver
// Metodologia.md §5.3.2) -- puede traer un prefijo de clave CSG pegado al
// inicio ("010.000.6153.00 PARACETAMOL..."). Esos dígitos son la CLAVE, no
// dosis/cantidad del medicamento, y ensuciarían el matching por conteo de
// números de abajo (falsos positivos por coincidencia numérica accidental
// contra la clave de OTRO producto en el catálogo/compendio). Se recortan
// con quitarPrefijoClave() (scripts/lib/dataset.js, mismo CLAVE_RE que usa
// extract.js) solo para esta comparación -- `producto` en sí se guarda crudo
// en el dataset, esto no lo toca.
function extraerNumerosProducto(producto) {
  return extraerNumeros(quitarPrefijoClave(producto));
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
// script (validacion_nombre_local, etc.), que ya pasaron por esta
// verificación y no hay que volver a derivar.
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
// describe. Si cualquiera de los dos lados (contrato o la propia ficha CUCoP)
// no trae ningún número, no hay nada que comparar -- se confía por defecto en
// vez de tratarlo como 0 coincidencias (que penalizaría fichas CUCoP sin
// dosis en su descripción, ej. "EMICIZUMAB" a secas, sin ser evidencia real
// de que el código esté mal citado).
//
// Recibe `numsProducto` ya calculado (ver main()) en vez de `producto` crudo
// -- si esto falla, buscarEnCompendioLocal() necesita el mismo cálculo para
// el MISMO registro; resolverlo una sola vez (memoizado en main(), no aquí)
// evita repetir quitarPrefijoClave()/normalizarTexto() dos veces.
function cucopEsValido(numsProducto, cveCucop, cucopMap) {
  const entry = cucopMap[cveCucop];
  if (!entry) return false;
  const numsEntry = extraerNumeros(entry.descripcion);
  if (numsProducto.length === 0 || numsEntry.length === 0) return true;
  return contarCoincidenciasNumericas(numsProducto, numsEntry) >= 2;
}

// Busca en el compendio local con el texto del PROPIO registro (ya no con el
// texto más largo de un grupo de productos "iguales") -- mismo umbral de alta
// confianza que el resto del script.
//
// Memoizado por el texto EXACTO de `producto`: al evaluar por registro en vez
// de por grupo canónico, este escaneo O(tamaño del compendio) pasa de
// correr una vez por producto único a una vez por registro sin resolver --
// muchas instituciones citan la MISMA descripción palabra por palabra, así
// que cachear por string exacto recupera casi todo el ahorro que daba el
// agrupamiento viejo, sin prestarle nada a un registro con texto distinto
// (es memoización de una función pura, no propagación entre registros).
// `obtenerNumsProducto` es un getter perezoso (no el valor ya calculado): en
// un cache hit no hace falta tocarlo, así que el chequeo de cache de arriba
// sigue evitando también el costo de extraerNumerosProducto(), no solo el
// del escaneo del compendio.
const cacheCompendioLocal = new Map();
function buscarEnCompendioLocal(producto, obtenerNumsProducto, compendio) {
  if (cacheCompendioLocal.has(producto)) return cacheCompendioLocal.get(producto);
  const numsProducto = obtenerNumsProducto();
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
  const resultado = mejorScore >= 2 ? { clave: mejorClave, score: mejorScore } : null;
  cacheCompendioLocal.set(producto, resultado);
  return resultado;
}

function asignar(data, i, clave, compendio, fuente, corregidos) {
  if (data[i].clave === clave && data[i].clave_fuente === fuente) return;
  corregidos.push({
    producto: data[i].producto.slice(0, 100),
    codigo_contrato: data[i].codigo_contrato,
    clave_anterior: data[i].clave,
    clave_nueva: clave,
    fuente,
  });
  data[i].clave = clave;
  data[i].grupo_terapeutico = clave && compendio[clave] ? compendio[clave].grupos_terapeuticos : null;
  data[i].clave_fuente = fuente;
}

async function main() {
  const data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
  const compendio = JSON.parse(fs.readFileSync(COMPENDIO_PATH, 'utf8'));
  const cucopMap = loadCucopMap();

  console.log(`Evaluando ${data.length} registros de forma individual (sin agrupar por producto)...`);

  const corregidos = [];
  const sinResolver = [];
  const contadores = { ya_confiables: 0, por_cucop_valido: 0, por_local: 0, vaciados_cucop_invalido: 0, sin_resolver: 0 };

  for (let i = 0; i < data.length; i++) {
    const registro = data[i];
    const fuente = inferirFuente(registro, cucopMap);

    if (esFuenteConfiable(fuente)) {
      contadores.ya_confiables++;
      continue;
    }

    const cucopEntry = registro.cve_cucop && cucopMap[registro.cve_cucop];
    const claveCucop = cucopEntry && cucopEntry.clave;
    // Getter perezoso y memoizado: cucopEsValido y buscarEnCompendioLocal
    // pueden necesitar el mismo cálculo para el MISMO `producto` (clave vía
    // CUCoP presente pero no autoverifica) -- se resuelve como máximo una
    // vez por registro, y nunca si buscarEnCompendioLocal tiene cache hit.
    let numsProducto;
    const obtenerNumsProducto = () => numsProducto ?? (numsProducto = extraerNumerosProducto(registro.producto));
    if (claveCucop && cucopEsValido(obtenerNumsProducto(), registro.cve_cucop, cucopMap)) {
      asignar(data, i, claveCucop, compendio, 'cucop', corregidos);
      contadores.por_cucop_valido++;
      continue;
    }

    const local = buscarEnCompendioLocal(registro.producto, obtenerNumsProducto, compendio);
    if (local) {
      asignar(data, i, local.clave, compendio, 'validacion_nombre_local', corregidos);
      contadores.por_local++;
      continue;
    }

    // Nada encontrado: si el registro tenía una clave vía CUCoP que no
    // autoverificó, se vacía explícitamente -- no dejar un valor que ya se
    // sabe sin evidencia solo porque no hay reemplazo mejor.
    if (fuente === 'cucop') {
      asignar(data, i, null, compendio, null, corregidos);
      contadores.vaciados_cucop_invalido++;
    }
    contadores.sin_resolver++;
    sinResolver.push({ producto: registro.producto.slice(0, 100), codigo_contrato: registro.codigo_contrato });
  }

  fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2), 'utf8');
  await guardarExcel(EXCEL_PATH, data);
  fs.writeFileSync(REPORTE_PATH, JSON.stringify({
    generado: new Date().toISOString(),
    resumen: { registros: data.length, registros_corregidos: corregidos.length, ...contadores },
    corregidos,
    sin_resolver: sinResolver,
  }, null, 2), 'utf8');

  console.log('\n--- Resumen ---');
  console.log(`Registros evaluados: ${data.length}`);
  console.log(`Ya confiables, sin tocar: ${contadores.ya_confiables}`);
  console.log(`Resueltos por CUCoP autoverificado (propio): ${contadores.por_cucop_valido}`);
  console.log(`Resueltos por compendio local (propio): ${contadores.por_local}`);
  console.log(`Sin resolver: ${contadores.sin_resolver} (de esos, ${contadores.vaciados_cucop_invalido} tenían clave vía CUCoP inválida que se vació)`);
  console.log(`Registros corregidos/actualizados en total: ${corregidos.length}`);
  console.log(`Reporte: ${REPORTE_PATH}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
