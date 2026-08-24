// Pipeline principal: filtra el CSV masivo de Compras MX por medicamentos (partida 25301),
// extrae precio unitario por contrato via Playwright, cruza contra el Compendio Nacional
// de Medicamentos, y guarda el resultado consolidado.
//
// Uso:
//   node scripts/extract.js data/raw/contratos_2026.csv [--limit N] [--out data/precios.json]

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const { parse } = require('csv-parse/sync');
const { chromium } = require('playwright');
const { CLAVE_RE, guardarExcel, dividirProductoSiAplica, cargarDatasetCompleto, guardarManifest } = require('./lib/dataset');

// Las transformaciones que este pipeline le hacía a `producto` (sustituir por
// ficha CUCoP+ cuando venía degenerado, recortar coletilla administrativa,
// recortar prefijos/viñetas, quitarle el prefijo de clave) se desactivaron a
// propósito el 2026-08-21 -- ver Metodologia.md §5.3.2/§5.3.3/§5.3.4 para el
// detalle completo (regex, casos reales, cifras) que queda documentado ahí
// como referencia para cuando se retomen. `producto` ahora es el texto crudo
// de la fuente, sin ningún recorte ni sustitución, para poder identificar
// todos los casos reales antes de diseñar cómo separar molécula/dosis/
// presentación (ver brainstorming pendiente).
const HASH_RE = /detalle\/([a-f0-9]+)\/procedimiento/i;
const NAV_TIMEOUT = 30000;
const CLICK_TIMEOUT = 15000;
const RESPONSE_TIMEOUT = 10000;
const DELAY_BETWEEN_CONTRATOS_MS = 400;
const DELAY_BETWEEN_EXPEDIENTES_MS = 800;
const CONTRATO_HARD_TIMEOUT_MS = 25000; // nunca dejar que un solo contrato cuelgue el pipeline (sube un poco vs. antes: ahora la búsqueda es exhaustiva y puede recorrer más páginas)
const MAX_INTENTOS = 3; // tras esto, un expediente con error se marca como fallo permanente y ya no se reintenta solo
// La tabla de contratos del expediente muestra 100 filas por página (confirmado
// en vivo el 2026-08-17). Por debajo de este umbral, un expediente cabe
// siempre en una sola página -- buscar contrato por contrato (findAndClickContrato)
// y la pasada única por página cuestan exactamente lo mismo, así que no vale
// la pena arriesgar el camino ya probado que corre sobre el 100% del CSV.
// Por arriba, findAndClickContrato reinicia a la página 1 en CADA búsqueda, y
// el orden de nuestra cola no tiene relación con el orden de la tabla del
// sitio (verificado: las filas de un mismo expediente en el CSV masivo están
// dispersas por todo el archivo, no contiguas) -- el costo de paginación
// escala con contratos × páginas en vez de solo páginas. Para el outlier de
// 3,221 contratos visto en vivo (~33 páginas a 100/página), eso es del orden
// de ~53,000 clicks de "siguiente" con el buscador actual vs. ~33 con la
// pasada única -- candidato fuerte a ser la causa principal de las corridas
// de 5+ horas por expediente grande. Si la pasada única demuestra funcionar
// bien en producción, la idea es unificar todo a ese único camino y retirar
// findAndClickContrato -- ver grill-me del 2026-08-17.
const UMBRAL_PASADA_UNICA = 100;
// Checkpoint por TIEMPO, no por cantidad de expedientes: las ventanas de
// proceso en background pueden cortarse tras procesar muy pocos expedientes
// (se observaron cortes con 0-7 expedientes procesados), así que hace falta
// guardar seguido -- pero un checkpoint cada 1 expediente (probado
// 2026-08-11) reescribe el JSON completo de resultados/errores en cada uno,
// y con la cola ordenada chicos-primero (la mayoría del CSV) eso son miles
// de reescrituras completas seguidas al arranque de la corrida, cada una más
// cara que la anterior según crece el array -- costo O(n²) que además
// bloquea el event loop justo en la fase que más expedientes procesa por
// segundo. Con un mínimo de tiempo entre escrituras en vez de entre
// expedientes, una ráfaga de expedientes chicos que terminan casi juntos
// colapsa en un solo checkpoint en vez de uno por cada uno, sin perder la
// garantía original (2s sigue siendo bastante margen por debajo de los
// cortes más agresivos observados). El checkpoint final de main() sigue
// garantizando que el último tramo, más corto que el intervalo, no se pierda.
const CHECKPOINT_MIN_INTERVALO_MS = 2000;

function withTimeout(promise, ms, label) {
  // Promise.race no cancela al perdedor -- sin clearTimeout, este timer queda
  // vivo hasta cumplirse `ms` aunque `promise` ya haya ganado la carrera.
  // Se llama una vez por contrato (mucho más seguido que el deadline de
  // grupo, que ya tiene el mismo fix), así que sin esto cada contrato deja
  // un timer huérfano de hasta CONTRATO_HARD_TIMEOUT_MS colgado.
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`hard_timeout: ${label}`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

function parseArgs(argv) {
  const args = { input: null, limit: Infinity, out: null, concurrency: 1, reset: false };
  const rest = argv.slice(2);
  args.input = rest.find(a => !a.startsWith('--'));
  const limitIdx = rest.indexOf('--limit');
  if (limitIdx !== -1) args.limit = parseInt(rest[limitIdx + 1], 10);
  const outIdx = rest.indexOf('--out');
  if (outIdx !== -1) args.out = rest[outIdx + 1];
  const concIdx = rest.indexOf('--concurrency');
  if (concIdx !== -1) args.concurrency = parseInt(rest[concIdx + 1], 10);
  args.reset = rest.includes('--reset');

  // Identifica de qué CSV viene una corrida -- se guarda como `origen` en
  // cada registro (ver buildRegistro). IMPORTANTE: el prefijo de
  // codigo_contrato (p. ej. "C-2025-...") NO sirve para esto -- es el año en
  // que el gobierno numeró el contrato, no el archivo de donde se scrapeó; un
  // contrato vigente multi-año puede tener prefijo de un año y venir
  // legítimamente del CSV de otro. `origen` es la única fuente de verdad
  // confiable sobre qué CSV produjo cada registro (usado por --reset abajo).
  if (args.input) {
    const m = path.basename(args.input).match(/(\d{4})/);
    args.origen = m ? m[1] : path.basename(args.input, path.extname(args.input));
  }
  // Un archivo por origen (docs/data.<año>.json), no un docs/data.json único
  // -- GitHub rechaza archivos >100 MB, y el combinado los superó apenas se
  // juntaron 2024+2025+2026 (2026-08-23). Ver scripts/lib/dataset.js para
  // cómo se descubren/mergean estos archivos en validar-claves.js y
  // audit-dataset.js.
  if (!args.out) args.out = path.join(__dirname, '..', 'docs', `data.${args.origen}.json`);
  return args;
}

function loadCompendio() {
  const p = path.join(__dirname, '..', 'data', 'compendio_medicamentos.json');
  return JSON.parse(fs.readFileSync(p, 'utf8'));
}

// Mapa cve_cucop -> clave CSG, precomputado una sola vez por
// scripts/build-cucop.js a partir del catálogo oficial CUCoP+
// (data/raw/cucop.xlsx) y commiteado como data/cucop_medicamentos.json --
// igual que data/compendio_medicamentos.json (§ loadCompendio). Evita
// parsear el xlsx crudo en cada corrida y, sobre todo, evita depender de que
// ese xlsx exista en el entorno de CI (no se versiona, ver .gitignore; solo
// el JSON derivado sí). Toda la verificación real (¿esta clave vía CUCoP es
// consistente con lo que el contrato describe? ¿coincide con otras instancias
// del mismo producto?) vive en scripts/validar-claves.js, que corre después
// sobre el dataset completo -- un solo lugar para toda la lógica de
// confiabilidad, en vez de repartirla entre extracción y validación.
function loadCucopMap() {
  const p = path.join(__dirname, '..', 'data', 'cucop_medicamentos.json');
  return JSON.parse(fs.readFileSync(p, 'utf8'));
}

function loadFilteredRows(csvPath, limit) {
  const raw = fs.readFileSync(csvPath, 'latin1'); // el CSV viene en latin1/windows-1252
  const records = parse(raw, { skip_empty_lines: true });
  const header = records[0];
  const idx = (name) => {
    const i = header.indexOf(name);
    if (i === -1) throw new Error(`Columna no encontrada en el CSV: "${name}"`);
    return i;
  };

  const col = {
    institucion: idx('Siglas de la Institución'),
    codigoExpediente: idx('Código del expediente'),
    partidaEspecifica: idx('Partida específica'),
    compraConsolidada: idx('Compra consolidada'),
    codigoContrato: idx('Código del contrato'),
    numContrato: idx('Núm. del contrato'),
    fechaFirmaContrato: idx('Fecha firma contrato'),
    fechaFallo: idx('Fecha de fallo'),
    fechaInicioContrato: idx('Fecha de inicio del contrato'),
    fechaFinContrato: idx('Fecha de fin del contrato'),
    tipoContrato: idx('Tipo de contrato'),
    // Hay dos columnas "Moneda" en el CSV (una junto a "Importe DRC", otra
    // junto a los montos mín/máx del bloque "Estatus Contrato" que no se usa
    // en este pipeline) -- idx() con header.indexOf toma la primera, la del
    // bloque DRC, que es la que viene siempre poblada (verificado: 0% vacía
    // en los 4 CSV, contra ~15-95% vacía en la segunda) y comparte bloque con
    // fechaInicioContrato/fechaFinContrato de abajo.
    moneda: idx('Moneda'),
    proveedor: idx('Proveedor o contratista'),
    numeroProcedimiento: header.indexOf('Número de procedimiento') !== -1 ? idx('Número de procedimiento') : idx('Número del procedimiento'),
    direccionAnuncio: idx('Dirección del anuncio'),
  };

  const rows = [];
  for (let i = 1; i < records.length; i++) {
    const r = records[i];
    if (!r || r.length < header.length) continue;
    const partida = r[col.partidaEspecifica] || '';
    if (!partida.includes('25301')) continue;
    const anuncio = r[col.direccionAnuncio] || '';
    const codigoContrato = r[col.codigoContrato] || '';
    if (!anuncio || !codigoContrato) continue;
    const hashMatch = anuncio.match(HASH_RE);
    if (!hashMatch) continue;

    rows.push({
      hash: hashMatch[1],
      codigoContrato,
      numContrato: r[col.numContrato] || '',
      institucion: r[col.institucion] || '',
      proveedor: r[col.proveedor] || '',
      procedimiento: r[col.numeroProcedimiento] || '',
      tipoProcedimiento: (r[col.compraConsolidada] || '').trim().toUpperCase() === 'SI' ? 'CONSOLIDADA' : 'NO CONSOLIDADA',
      fechaFirmaContrato: r[col.fechaFirmaContrato] || '',
      fechaFallo: r[col.fechaFallo] || '',
      fechaInicioContrato: r[col.fechaInicioContrato] || '',
      fechaFinContrato: r[col.fechaFinContrato] || '',
      moneda: r[col.moneda] || '',
      // Valor oficial de la fuente ("Abierto"/"Cerrado") -- se captura tal
      // cual, no se infiere de si cantidad_maxima viene poblada, para que
      // quede una sola fuente de verdad sin margen de inconsistencia entre
      // ambas señales.
      tipoContrato: r[col.tipoContrato] || '',
      codigoExpediente: r[col.codigoExpediente] || '',
    });
    if (rows.length >= limit) break;
  }
  return rows;
}

function groupByHash(rows) {
  const groups = new Map();
  for (const row of rows) {
    if (!groups.has(row.hash)) groups.set(row.hash, []);
    groups.get(row.hash).push(row);
  }
  return groups;
}

// Localizadores del paginador de la tabla "DATOS RELEVANTES DEL CONTRATO",
// compartidos entre findAndClickContrato (búsqueda por contrato) y
// procesarPorPaginado (pasada única, ver UMBRAL_PASADA_UNICA) -- ambos
// necesitan exactamente los mismos tres. Los botones se acotan al que sigue
// a ese encabezado en el DOM porque la página puede tener varias tablas
// paginadas (anexos, requerimientos, contratos) con la misma clase.
function localizadoresTablaContratos(page) {
  const scope = page.locator('text=DATOS RELEVANTES DEL CONTRATO');
  return {
    firstBtn: scope.locator('xpath=following::*[contains(@class,"p-paginator-first")][1]'),
    nextBtn: scope.locator('xpath=following::*[contains(@class,"p-paginator-next")][1]'),
    filas: scope.locator('xpath=following::td[contains(@class,"p-link2")]'),
  };
}

async function irAPrimeraPagina(page, firstBtn) {
  const firstDisabled = await firstBtn.evaluate(el => el.classList.contains('p-disabled')).catch(() => true);
  if (!firstDisabled) {
    await firstBtn.click({ timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(600);
  }
}

// Devuelve false si ya estaba en la última página (nada a dónde avanzar).
async function avanzarPagina(page, nextBtn) {
  const disabled = await nextBtn.evaluate(el => el.classList.contains('p-disabled')).catch(() => true);
  if (disabled) return false;
  await nextBtn.click({ timeout: 5000 }).catch(() => {});
  await page.waitForTimeout(600);
  return true;
}

async function findAndClickContrato(page, numContrato) {
  const { firstBtn, nextBtn } = localizadoresTablaContratos(page);

  // Reiniciar siempre a la primera página: como esta función se llama una vez
  // por cada contrato objetivo dentro del mismo expediente, sin este reset la
  // búsqueda del segundo contrato en adelante arrancaría desde donde quedó la
  // paginación tras buscar el anterior, en vez de desde el principio.
  await irAPrimeraPagina(page, firstBtn);

  // Búsqueda exhaustiva: recorre todas las páginas hasta encontrar el contrato
  // o agotar la tabla. Antes el único límite era el timeout duro por contrato
  // (CONTRATO_HARD_TIMEOUT_MS) que envuelve a esta función -- pero un
  // Promise.race no cancela el trabajo en curso: si ese timeout externo se
  // dispara, este bucle seguía corriendo "huérfano" contra la misma `page`
  // (verificado en vivo: un caso real donde nextBtn nunca reportó disabled y
  // el proceso quedó colgado indefinidamente). Ahora el propio bucle respeta
  // un límite de tiempo, con margen bajo CONTRATO_HARD_TIMEOUT_MS para que
  // termine solo antes de que el timeout externo lo intente cortar.
  const deadline = Date.now() + CONTRATO_HARD_TIMEOUT_MS - 3000;
  while (Date.now() < deadline) {
    const link = page.locator('td.p-link2', { hasText: numContrato }).first();
    if (await link.count() > 0 && await link.isVisible().catch(() => false)) {
      await link.click();
      return true;
    }
    if (!(await avanzarPagina(page, nextBtn))) return false; // se recorrieron todas las páginas, no está en la tabla
  }
  return false; // se agotó el tiempo de búsqueda sin encontrar el contrato
}

// Clickear una fila abre el detalle del contrato en un modal (p-dialog) que
// queda flotando SOBRE la tabla de contratos -- nunca se cierra solo. Si no se
// cierra aquí, la búsqueda del siguiente contrato "encuentra" su fila (sigue
// existiendo en el DOM, Playwright la reporta visible) pero el click nunca
// aterriza porque el overlay del modal intercepta el evento de puntero, y
// Playwright reintenta el click hasta agotar su propio timeout -- lo cual
// arrastra a TODOS los contratos siguientes del mismo expediente a fallar en
// cascada. Se verificó en vivo: sin este cierre, un expediente de 22
// contratos terminaba con 21 en hard_timeout (todos menos el primero).
async function closeModalIfOpen(page) {
  const closeBtn = page.locator('.p-dialog-header-close').first();
  if (await closeBtn.count() > 0 && await closeBtn.isVisible().catch(() => false)) {
    await closeBtn.click({ timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(400);
  }
}

// `numContrato` (buscado por texto en findAndClickContrato) no es único --
// dos contratos del mismo expediente pueden mostrar el mismo texto visible
// (confirmado en vivo el 2026-08-22: "I-97-2025-1" en dos filas distintas,
// codigoContrato C-2026-00000266 y C-2026-00000294). Antes esto se
// diagnosticaba mal: el listener solo aceptaba la URL que coincidiera
// EXACTO con `contrato.codigoContrato`, así que un click que aterrizaba en
// el contrato equivocado nunca disparaba ese filtro -- `captured` quedaba
// `null` para siempre y el contrato correcto salía como
// 'timeout_esperando_respuesta' (parecía que el sitio no respondía, cuando
// en realidad sí respondió, pero para otro contrato). Ahora se captura
// CUALQUIER respuesta de detallepartidas para este `hash` y se lee el
// codigoContrato real de la URL (mismo patrón que ya usa la pasada única en
// procesarPorPaginado) -- si no coincide con el que se buscaba, se reporta
// como 'numero_contrato_ambiguo', un diagnóstico preciso en vez de un
// timeout genérico. No se intenta resolver la ambigüedad automáticamente
// acá (a diferencia de la pasada única, que sí puede porque recorre TODAS
// las filas): es un caso raro (2 de ~17,000 contratos en la corrida de
// 2026-08-22) y `findAndClickContrato` solo sabe buscar por texto, no tiene
// forma de distinguir cuál de las filas duplicadas es la correcta.
const DETALLE_RE_CONTRATO = (hash) => new RegExp(`/whitney/sitiopublico/detallepartidas/${hash}/([^/?]+)`);

async function extractContrato(page, hash, contrato) {
  let captured = null;
  let codigoRecibido = null;
  const detalleRe = DETALLE_RE_CONTRATO(hash);
  const onResponse = async (res) => {
    const m = res.url().match(detalleRe);
    if (m) {
      codigoRecibido = decodeURIComponent(m[1]);
      try { captured = JSON.parse(await res.text()); } catch (e) { /* ignore */ }
    }
  };
  page.on('response', onResponse);
  try {
    const found = await findAndClickContrato(page, contrato.numContrato);
    if (!found) return { error: 'contrato_no_encontrado_en_tabla' };

    const start = Date.now();
    while (!captured && Date.now() - start < RESPONSE_TIMEOUT) {
      await page.waitForTimeout(200);
    }
    if (!captured) return { error: 'timeout_esperando_respuesta' };
    if (codigoRecibido !== contrato.codigoContrato) {
      return { error: 'numero_contrato_ambiguo', detalle: `se buscaba ${contrato.codigoContrato} (num "${contrato.numContrato}"), llegó respuesta de ${codigoRecibido}` };
    }
    if (!captured.success) return { error: 'respuesta_sin_exito', detalle: captured };

    return { data: captured.data || [] };
  } finally {
    page.off('response', onResponse);
    await closeModalIfOpen(page);
  }
}

// Fallback cuando extractContrato reporta 'numero_contrato_ambiguo' -- en vez
// de darse por vencido, prueba TODAS las filas que compartan ese numContrato
// (no solo la primera que encuentra Playwright), leyendo el codigoContrato
// real de cada respuesta hasta dar con la que buscábamos. Mismo mecanismo que
// procesarPorPaginado (más abajo), pero acotado a un solo contrato en vez de
// recorrer todo el expediente -- solo tiene sentido llamarlo para expedientes
// por debajo de UMBRAL_PASADA_UNICA (los grandes ya se resuelven bien con
// procesarPorPaginado). Confirmado en vivo el 2026-08-23: de 18 casos
// ambiguos vistos en expedientes chicos, el patrón es 100% determinista --
// `findAndClickContrato` siempre aterriza en la misma fila física, así que 3
// reintentos con el camino viejo nunca lo resolvían solos.
async function resolverAmbiguoPorPaginado(page, hash, contrato) {
  const detalleRe = DETALLE_RE_CONTRATO(hash);
  const normalizarEspacios = (s) => String(s || '').replace(/\s+/g, ' ').trim();
  const numContratoBuscado = normalizarEspacios(contrato.numContrato);
  const { firstBtn, nextBtn, filas } = localizadoresTablaContratos(page);

  const abrirYCapturar = async (i) => {
    let captured = null;
    const onResponse = async (res) => {
      const m = res.url().match(detalleRe);
      if (m) {
        try { captured = { codigo: decodeURIComponent(m[1]), json: JSON.parse(await res.text()) }; }
        catch (e) { /* respuesta no-JSON, se ignora -- captured sigue null */ }
      }
    };
    page.on('response', onResponse);
    try {
      await filas.nth(i).click({ timeout: CLICK_TIMEOUT }).catch(() => {});
      const start = Date.now();
      while (!captured && Date.now() - start < RESPONSE_TIMEOUT) await page.waitForTimeout(200);
      return captured;
    } finally {
      page.off('response', onResponse);
      await closeModalIfOpen(page);
    }
  };

  await irAPrimeraPagina(page, firstBtn);
  // Expedientes ≤ UMBRAL_PASADA_UNICA (100 contratos) caben en 1-2 páginas
  // (100 filas/página confirmado en vivo) -- este tope es solo una cota de
  // seguridad contra un paginador atascado, no un límite real esperado.
  const maxPaginas = 5;
  for (let paginasVisitadas = 0; paginasVisitadas < maxPaginas; paginasVisitadas++) {
    const textos = (await filas.allInnerTexts().catch(() => [])).map(normalizarEspacios);
    for (let i = 0; i < textos.length; i++) {
      if (textos[i] !== numContratoBuscado) continue;
      const captured = await abrirYCapturar(i);
      if (!captured || captured.codigo !== contrato.codigoContrato) continue; // no era la fila que buscábamos, seguir probando
      return captured.json.success ? { data: captured.json.data || [] } : { error: 'respuesta_sin_exito', detalle: captured.json };
    }
    if (!(await avanzarPagina(page, nextBtn))) break;
  }
  return { error: 'numero_contrato_ambiguo_sin_resolver', detalle: `ninguna fila con num "${contrato.numContrato}" respondió con ${contrato.codigoContrato}` };
}

function buildRegistro(contrato, item, compendio, cucopMap, stats, origen) {
  // Todo lo que llega aquí ya viene de un contrato con Partida específica 25301
  // (medicamentos), así que se conserva aunque no traiga la clave del compendio
  // embebida en la descripción -- solo algunas instituciones (grandes/federales)
  // la citan; el resto no. Cuando SÍ hay clave y existe en el compendio, se
  // enriquece con grupo_terapéutico; si no, quedan en null (baja confianza pero
  // se conserva la fila).
  // `producto` es el texto crudo de la fuente, tal cual -- ver nota arriba de
  // HASH_RE sobre por qué se desactivaron las transformaciones que antes lo
  // recortaban/sustituían. `m` se sigue usando abajo solo para `clave`, no
  // para tocar `producto`.
  const m = String(item.descripcion || '').match(CLAVE_RE);
  const producto = String(item.descripcion || '').trim();

  // Prioridad para asignar clave: 1) extraerla directamente de la descripción
  // de la institución (más confiable); 2) si no vino, recuperarla vía el
  // catálogo CUCoP. Se guarda clave_fuente para que scripts/validar-claves.js
  // sepa qué corrida completa se encargue de verificar/corregir -- la
  // extracción no se detiene a autoverificar, solo agarra lo que encuentre
  // (ver Metodologia.md §7/§9.1 sobre por qué la vía CUCoP puede venir mal).
  let clave = m ? `${m[1]}.${m[2]}.${m[3]}.${m[4]}` : null;
  let claveFuente = clave ? 'descripcion' : null;
  if (!clave) {
    const viaCucop = cucopMap[item.cve_cucop];
    // viaCucop.clave puede ser null (entrada del catálogo sin prefijo de
    // clave CSG en su descripción -- ver scripts/build-cucop.js) -- ahí no hay clave
    // que adoptar, solo descripción (ya usada arriba para `producto`).
    if (viaCucop && viaCucop.clave) {
      clave = viaCucop.clave;
      claveFuente = 'cucop';
    }
  }

  const compEntry = clave ? compendio[clave] : null;

  let cantidad = item.cantidad_maxima ?? item.cantidad ?? item.cantidad_minima ?? null;

  // El campo "precio_unitario" que entrega la API a veces viene mal cargado en
  // el origen (se detectó un caso real donde traía copiado el valor de
  // "subtotal_cant_max" en vez del precio real -- $140,605,012/unidad en una
  // vacuna, un error del propio sistema de Compras MX). Como el subtotal SÍ
  // es consistente con cantidad_minima en todos los casos verificados
  // (contratos abiertos y de cantidad fija), se recalcula el precio dividiendo
  // subtotal / cantidad_minima en vez de confiar ciegamente en precio_unitario.
  let precioUnitario = item.precio_unitario ?? null;
  if (item.subtotal != null && item.cantidad_minima) {
    const precioCalculado = item.subtotal / item.cantidad_minima;
    if (precioUnitario == null || Math.abs(precioCalculado - precioUnitario) / precioCalculado > 0.01) {
      if (precioUnitario != null) {
        stats.preciosCorregidos.push({ codigo_contrato: contrato.codigoContrato, cve_cucop: item.cve_cucop, precio_api: precioUnitario, precio_calculado: +precioCalculado.toFixed(2) });
      }
      precioUnitario = precioCalculado;
    }
  }

  // Contratos "por monto" (tipo_contrato_abierto: "MONTO" en la fuente): el
  // compromiso es un techo de gasto, no una cantidad, así que la API nunca
  // trae cantidad/cantidad_minima/cantidad_maxima para estos ítems -- vienen
  // vacíos desde el origen, no es un hueco de la extracción. Sí trae
  // `subtotal` (el "Monto de la Oferta" que se ve en el detalle del
  // contrato), que junto con precio_unitario permite derivar la cantidad
  // implícita de la oferta: subtotal = precio_unitario × cantidad.
  if (cantidad == null && item.subtotal != null && precioUnitario) {
    cantidad = Math.round(item.subtotal / precioUnitario);
    stats.cantidadesDerivadas.push({ codigo_contrato: contrato.codigoContrato, cve_cucop: item.cve_cucop, subtotal: item.subtotal, precio_unitario: precioUnitario, cantidad_derivada: cantidad });
  }

  // `subtotal` es el monto total ya calculado por la fuente para el ítem --
  // cuando no hay cantidad_minima/cantidad_maxima distintas que preservar
  // (Cerrado, o "por monto" sin ningún rango), es el número más confiable
  // posible para valor_minimo/valor_maximo, más confiable incluso que
  // recalcularlo multiplicando por `cantidad`. Esto importa en particular
  // para el caso "por monto": ahí `cantidad` ya se derivó arriba dividiendo
  // subtotal/precio_unitario, así que multiplicarla de vuelta por
  // precio_unitario sería un viaje de ida y vuelta innecesario (subtotal →
  // cantidad → subtotal) que además puede perder precisión por el redondeo
  // intermedio de `cantidad` -- si ya se tiene subtotal, se usa directo.
  const valorDirecto = item.subtotal != null ? +item.subtotal.toFixed(2) : null;

  // cantidad_minima/cantidad_maxima quedan SIEMPRE pobladas (nunca null): si
  // el ítem no trae un rango genuino de origen (Cerrado, o "por monto"),
  // AMBAS colapsan al mismo valor -- `cantidad`, la única cifra disponible en
  // ese caso -- en vez de dejarlas en null o (peor) dejar que cada una caiga
  // a un respaldo distinto. La API puede traer `item.cantidad_minima` poblado
  // con un número que no guarda relación con el resto del ítem incluso cuando
  // `item.cantidad_maxima` viene vacío (visto en casos reales) -- si cada
  // campo cayera a su propio `?? cantidad` por separado, cantidad_minima
  // podría quedar en ese número suelto mientras cantidad_maxima cae a
  // `cantidad`, produciendo un "rango" falso (min != max) que no tiene
  // ninguna relación con valor_minimo/valor_maximo (que si son iguales entre
  // sí en este caso). Por eso ambas dependen del MISMO `hayRangoGenuino` que
  // decide valor_minimo/valor_maximo abajo, no de su propio `??` independiente.
  //
  // Mismo criterio para valor_minimo/valor_maximo: piso y techo de exposición
  // contractual, bien definidos para los dos tipos de contrato sin necesidad
  // de segmentar antes de calcular. En "Cerrado" caen al mismo número (no hay
  // rango que representar); en "Abierto" divergen. Esto permite sumar/agregar
  // valor entre ambos tipos sin mezclar un "techo posible" (Abierto) con un
  // "monto real" (Cerrado) como si fueran la misma clase de número --
  // cualquier análisis agregado de gasto debe usar valor_minimo (piso
  // garantizado) o valor_maximo (techo de exposición) explícitamente
  // (Metodologia.md §5.5).
  //
  // Solo cuando SÍ hay un rango genuino (cantidad_minima y cantidad_maxima
  // distintas, "Abierto" real) hace falta calcular piso/techo por separado
  // multiplicando -- ahí `subtotal` no sirve como respaldo porque la API lo
  // reporta consistente con cantidad_minima únicamente (ver corrección de
  // precio_unitario arriba), así que usarlo para el techo estaría mal.
  const hayRangoGenuino = item.cantidad_minima != null && item.cantidad_maxima != null && item.cantidad_minima !== item.cantidad_maxima;
  const cantidadMinimaEfectiva = hayRangoGenuino ? item.cantidad_minima : cantidad;
  const cantidadMaximaEfectiva = hayRangoGenuino ? item.cantidad_maxima : cantidad;

  const valorMinimoCalculado = (precioUnitario != null && cantidadMinimaEfectiva != null) ? +(precioUnitario * cantidadMinimaEfectiva).toFixed(2) : null;
  const valorMaximoCalculado = (precioUnitario != null && cantidadMaximaEfectiva != null) ? +(precioUnitario * cantidadMaximaEfectiva).toFixed(2) : null;

  const valorMinimo = (!hayRangoGenuino && valorDirecto != null) ? valorDirecto : (valorMinimoCalculado ?? valorDirecto);
  const valorMaximo = (!hayRangoGenuino && valorDirecto != null) ? valorDirecto : (valorMaximoCalculado ?? valorDirecto);

  return {
    codigo_contrato: contrato.codigoContrato,
    // De qué CSV vino este registro (p. ej. "2025", "2026") -- NO se puede
    // derivar del prefijo de codigo_contrato, ver nota en parseArgs(). Es la
    // clave que usa --reset para poder borrar/reintentar un origen sin
    // arriesgar los demás.
    origen,
    num_contrato: contrato.numContrato,
    clave,
    clave_fuente: claveFuente,
    cve_cucop: item.cve_cucop || null,
    producto,
    tipo_insumo: 'MEDICAMENTO',
    grupo_terapeutico: compEntry ? compEntry.grupos_terapeuticos : null,
    procedimiento: contrato.procedimiento,
    tipo_procedimiento: contrato.tipoProcedimiento,
    // "Abierto"/"Cerrado" tal cual lo declara la fuente -- distingue de forma
    // explícita e inequívoca el esquema de cantidad de este registro (rango
    // mín/máx vs. cantidad fija), en vez de tener que inferirlo de si
    // cantidad_maxima viene poblada. Permite filtrar/agrupar por tipo sin
    // ambigüedad al integrar ambos esquemas en la misma tabla.
    tipo_contrato: contrato.tipoContrato || null,
    proveedor: contrato.proveedor,
    institucion: contrato.institucion,
    moneda: contrato.moneda || null,
    unidad_medida: item.um || null,
    precio_unitario: precioUnitario,
    cantidad_minima: cantidadMinimaEfectiva,
    cantidad_maxima: cantidadMaximaEfectiva,
    valor_minimo: valorMinimo,
    valor_maximo: valorMaximo,
    fecha_firma_contrato: contrato.fechaFirmaContrato,
    fecha_fallo: contrato.fechaFallo || null,
    fecha_inicio_contrato: contrato.fechaInicioContrato || null,
    fecha_fin_contrato: contrato.fechaFinContrato || null,
    // `item_id` (uniendo con docs/data.raw.json, ver procesarItems) se
    // agrega ahí, no acá -- buildRegistro no sabe en qué posición del array
    // de la respuesta vino este item.
  };
}

// `pendientes` es lo que realmente se procesa esta corrida (puede ser un
// subconjunto chico de un expediente grande ya parcialmente resuelto);
// `total` es el tamaño real del expediente. Se muestran juntos solo cuando
// difieren, para no ensuciar el caso común (expediente sin progreso previo).
function fmtContratos(pendientes, total) {
  return pendientes === total ? `${pendientes} contrato(s)` : `${pendientes}/${total} contrato(s) pendientes`;
}

async function processExpedienteGroup(browser, hash, contratos, compendio, cucopMap, ctx, origen, totalExpediente) {
  const url = `https://comprasmx.buengobierno.gob.mx/sitiopublico/#/sitiopublico/detalle/${hash}/procedimiento`;
  const resultados = [];
  const errores = [];
  // `buildRegistro` recibe `ctx.stats` directo (no un acumulador local) --
  // mismo motivo y mismo patrón que agregarResultado()/agregarError() abajo:
  // antes `stats` se mutaba en un objeto LOCAL al grupo, fundido al `stats`
  // de toda la corrida recién cuando processExpedienteGroup retornaba, así
  // que un registro huérfano post-deadline que resolvía tarde (ver
  // Promise.race de abajo) podía aportar a `stats` DESPUÉS de ese merge
  // puntual y perderse -- inconsistente con agregarResultado()/agregarError(),
  // que ya empujan en vivo al acumulador compartido y nunca pierden nada por
  // timing. Con `ctx.stats` mutado en vivo, ambos quedan con la misma
  // garantía: lo que se produce, se persiste, sin importar cuándo llegue.

  // Checkpoint por contrato: además de acumular en los arrays LOCALES de
  // este grupo (para los cómputos internos de abajo: yaResueltos/yaConError,
  // el resumen final), cada resultado/error se empuja de inmediato al
  // acumulador COMPARTIDO de toda la corrida (`ctx.allResultados`/
  // `ctx.allErrores`) -- antes, el checkpoint de un expediente grande recién
  // se escribía cuando el grupo COMPLETO terminaba (worker() fundía los
  // arrays locales al final); un expediente de miles de contratos (se vieron
  // hasta 3,221 en vivo) tardando horas significaba que cortar el proceso a
  // mitad de camino perdía TODO su progreso, sin importar cuántos contratos
  // ya se hubieran scrapeado con éxito. Con esto, `checkpointSiToca` (llamado
  // periódicamente más abajo, dentro del loop) puede escribir a disco
  // progreso real de un grupo que todavía sigue en curso.
  // Si un contrato ya había quedado marcado con el error sintético
  // 'grupo_timeout' (ver más abajo) y RECIÉN DESPUÉS llega su resultado o
  // error real (huérfano que resolvió más tarde que la ventana de gracia de
  // 2s), se descarta el sintético en vez de dejarlo duplicado -- sin esto,
  // el contrato quedaba contado dos veces (una vez como éxito/error real,
  // otra como 'grupo_timeout'), inflando intentosPorContrato de cara a la
  // próxima corrida sin que haya habido dos fallos genuinos.
  const descartarTimeoutSintetico = (codigoContrato) => {
    const esSintetico = (e) => e.codigoContrato === codigoContrato && e.error === 'grupo_timeout';
    const iLocal = errores.findIndex(esSintetico);
    if (iLocal !== -1) errores.splice(iLocal, 1);
    const iShared = ctx.allErrores.findIndex(esSintetico);
    if (iShared !== -1) ctx.allErrores.splice(iShared, 1);
  };
  const agregarResultado = (r) => {
    descartarTimeoutSintetico(r.codigo_contrato);
    resultados.push(r);
    ctx.allResultados.push(r);
  };
  // `ctx.allErrores` (el acumulador compartido, persistido a
  // docs/data.errores.json) guarda UNA sola entrada por contrato con un
  // contador `intentos`, no una copia por cada vez que falla -- antes cada
  // llamada empujaba una entrada nueva sin tocar las anteriores, así que un
  // contrato que fallaba 3 veces (MAX_INTENTOS) terminaba con 3 copias casi
  // idénticas en el archivo publicado, sin ninguna limpieza entre corridas
  // (necesario desde 2026-08-21 para que intentosPorContrato pueda contar
  // entre corridas -- ver comentario de más arriba sobre por qué se dejó de
  // descartar errores de contratos pendientes). `errores` (el array LOCAL de
  // este grupo, solo para el resumen impreso de ESTA corrida) sigue
  // agregando una entrada por evento -- no se persiste tal cual, no hace
  // falta consolidarlo.
  const agregarError = (e) => {
    if (e.error !== 'grupo_timeout') descartarTimeoutSintetico(e.codigoContrato);
    errores.push(e);
    const existente = ctx.allErrores.find(x => x.hash === e.hash && x.codigoContrato === e.codigoContrato);
    if (existente) {
      Object.assign(existente, e, { intentos: (existente.intentos || 1) + 1 });
    } else {
      ctx.allErrores.push({ ...e, intentos: 1 });
    }
  };

  // Compartido entre el camino de búsqueda por contrato y la pasada única
  // por página (ver UMBRAL_PASADA_UNICA) -- ambos terminan con la misma
  // respuesta cruda de detallepartidas, solo cambia CÓMO se llega a ella.
  // Devuelve cuántos registros agregó -- un contrato cuyos ítems son TODOS de
  // otra partida (filtro de abajo) legítimamente produce 0. Sin marcar eso de
  // alguna forma, el contrato nunca entra a `codigosConResultado` NI a
  // `errores` -- queda "pendiente" para siempre, y cada corrida futura lo
  // vuelve a abrir sin que el conteo de pendientes baje nunca (confirmado en
  // vivo el 2026-08-21: 137 expedientes de 2026 se reprocesaban en cada
  // intento de run-extract.sh sin converger). El llamador usa este valor para
  // registrar un error explícito cuando da 0, así el contrato entra al mismo
  // conteo de MAX_INTENTOS que cualquier otro fallo y eventualmente se marca
  // "fallo permanente" en vez de reprocesarse indefinidamente.
  const procesarItems = (contrato, data) => {
    let agregados = 0;
    data.forEach((item, idx) => {
      // El filtro de la partida 25301 en loadFilteredRows() opera a nivel
      // de CONTRATO (una fila del CSV masivo), no de ítem: un contrato con
      // al menos una línea de partida 25301 pasa el filtro, pero
      // detallepartidas() devuelve TODAS las líneas del contrato, incluidas
      // las de otras partidas (radiofármacos 25401, material de curación
      // 25501, etc.) si la compra fue mixta. Sin este segundo filtro por
      // ítem, esas líneas se colaban al dataset marcadas como MEDICAMENTO
      // (~32% del dataset en la corrida de 2026-08-09, detectado por cve_cucop
      // fuera de 25301-*).
      if (!item.cve_cucop || !String(item.cve_cucop).startsWith('25301')) return;
      // `item` crudo, sin transformar, a un archivo APARTE (docs/data.raw.json,
      // ver ctx.allRawItems) -- no en el registro mismo, para no inflar el
      // JSON que el dashboard público descarga entero en cada visita
      // (~41 MB extra estimados sobre el dataset completo si viviera ahí).
      // `item_id` (índice del item dentro de la respuesta de ESTE contrato)
      // es la llave para unir un registro de docs/data.json con su item
      // crudo acá -- necesaria porque un mismo codigo_contrato puede traer
      // varios ítems (varias líneas de partida 25301 en un solo contrato).
      const itemId = `${contrato.codigoContrato}#${idx}`;
      ctx.allRawItems.push({ item_id: itemId, codigo_contrato: contrato.codigoContrato, origen, item });
      const registro = buildRegistro(contrato, item, compendio, cucopMap, ctx.stats, origen);
      if (!registro) return;
      registro.item_id = itemId;
      // Casos puntuales confirmados a mano (ver PRODUCTOS_A_DIVIDIR en
      // lib/dataset.js) donde `producto` en realidad lista varios
      // ingredientes separados por " - " -- cada uno se guarda como su
      // propio registro, con el resto de los datos (clave, precio,
      // cantidad, item_id) idéntico. Para el caso común (no está en la
      // lista) esto devuelve [producto] sin cambios, un solo registro.
      for (const producto of dividirProductoSiAplica(registro.producto)) {
        agregarResultado(producto === registro.producto ? registro : { ...registro, producto });
        agregados++;
      }
    });
    return agregados;
  };

  // Pestaña nueva por expediente: reutilizar una sola pestaña entre navegaciones
  // de rutas con hash de esta SPA (Angular) resultó no confiable en pruebas.
  let page = await browser.newPage();

  let loaded = false;
  for (let intento = 0; intento < 2 && !loaded; intento++) {
    try {
      const waitDetalle = page.waitForResponse(
        res => res.url().includes(`/whitney/sitiopublico/expedientes/${hash}?id_proceso=procedimiento`) && res.status() === 200,
        { timeout: NAV_TIMEOUT }
      );
      // Si page.goto falla antes de que waitDetalle resuelva, esta promesa
      // queda abandonada; sin este catch, un rechazo tardío (timeout) se
      // vuelve una unhandled rejection que tumba TODO el proceso Node, no
      // solo este intento -- verificado en vivo el 2026-08-10.
      waitDetalle.catch(() => {});
      await page.goto(url, { waitUntil: 'load', timeout: NAV_TIMEOUT });
      await waitDetalle;
      await page.waitForTimeout(800); // margen para que Angular renderice la tabla tras la respuesta
      loaded = true;
    } catch (e) {
      // Un error POR CONTRATO, no uno solo a nivel expediente: calcularPendientes
      // indexa reintentos por hash|codigoContrato, así que un error sin
      // codigoContrato nunca hace match con ningún contrato real -- el
      // expediente jamás llegaba a "fallo permanente" y se reintentaba para
      // siempre en cada corrida, sin importar cuántas veces fallara la
      // navegación (verificado en vivo el 2026-08-10).
      if (intento === 1) {
        for (const contrato of contratos) {
          agregarError({ hash, codigoContrato: contrato.codigoContrato, origen, error: 'navegacion_fallida', detalle: String(e) });
        }
      }
    }
  }

  if (loaded) {
    // Límite global del expediente: cada contrato ya tiene su propio timeout
    // duro (CONTRATO_HARD_TIMEOUT_MS), pero withTimeout() usa Promise.race,
    // que NO cancela el trabajo huérfano que sigue corriendo contra la misma
    // `page` compartida después de que el timeout externo "gana" la carrera
    // -- el mismo patrón ya documentado arriba para findAndClickContrato,
    // pero sin cubrir el caso de que ese trabajo huérfano deje la page en un
    // estado que traba a TODOS los contratos siguientes del mismo worker.
    // Visto en vivo: un worker completo sin avanzar por más de una hora
    // (verificado 2026-08-11, corrida de 2025). Esta capa adicional fuerza el
    // cierre de la page (aborta cualquier operación colgada en curso) si el
    // expediente entero se pasa de un tiempo generoso. Los contratos que no
    // se alcanzaron a procesar SÍ quedan registrados como error ahora (ver
    // bloque tras el Promise.race de abajo) -- antes no contaban para nada,
    // así que un expediente crónicamente atorado (no un corte de ventana
    // cualquiera, sino la misma page rota una y otra vez) nunca llegaba a
    // "fallo permanente" y quemaba este mismo deadline generoso en cada
    // corrida, para siempre. `grupoDeadlineMs` ya da margen de sobra para
    // cualquier expediente sano (contratos.length * CONTRATO_HARD_TIMEOUT_MS,
    // el tiempo máximo que ya le tocaría a cada contrato individualmente) --
    // llegar a este punto es señal real de que algo está roto, no solo de
    // que el expediente es grande.
    const grupoDeadlineMs = Math.max(60000, contratos.length * CONTRATO_HARD_TIMEOUT_MS);
    const grupoStart = Date.now();

    // Pasada única por página: solo para expedientes grandes (ver
    // UMBRAL_PASADA_UNICA) -- recorre la tabla de página en página UNA sola
    // vez, en vez de reiniciar a la página 1 en cada búsqueda de contrato
    // (findAndClickContrato). numContrato NO está garantizado único, así que
    // no se decide de antemano qué fila es cuál contrato -- se clickea la
    // fila que matchea por texto, y el codigoContrato REAL se lee de la URL
    // de la respuesta capturada (mismo mecanismo que extractContrato, pero
    // sin fijar de antemano cuál codigoContrato se espera). Si la fila
    // resuelta no es ninguno de nuestros pendientes (numContrato compartido
    // con un contrato ajeno a este grupo, o ya resuelto), se descarta sin
    // contar como progreso ni como error -- no es parte de lo que este grupo
    // necesita. Definida DENTRO de processExpedienteGroup (no aparte, con
    // `page` como parámetro) a propósito: así sigue cerrando sobre la MISMA
    // variable `page` que el resto de la función, y el `if (!page) break`
    // de abajo reacciona igual al deadline de grupo que el camino chico.
    async function procesarPorPaginado() {
      const detalleRe = new RegExp(`/whitney/sitiopublico/detallepartidas/${hash}/([^/?]+)`);
      const porCodigo = new Map(contratos.map(c => [c.codigoContrato, c]));
      // El `td.p-link2` de la fila puede traer espacios internos/dobles o no
      // separables (nbsp) que .trim() solo no limpia -- la búsqueda vieja
      // (findAndClickContrato, vía `hasText`) tolera esto porque Playwright
      // normaliza espacios en `hasText`; acá se compara texto exacto, así
      // que hay que normalizar los dos lados (el numContrato del CSV Y el
      // texto de la fila) con el mismo criterio, o un contrato presente en
      // la tabla podría reportarse como "no encontrado" por una diferencia
      // de espacios que no tiene nada que ver con si existe o no.
      const normalizarEspacios = (s) => String(s || '').replace(/\s+/g, ' ').trim();
      // Set estático, nunca se achica: solo sirve para saltar de una fila
      // que a simple vista no es ninguno de nuestros pendientes, sin
      // clickearla. La corrección real ya la garantiza `porCodigo.has(...)`
      // después de capturar la respuesta -- si dos filas comparten texto y
      // ya resolvimos una, la otra igual pasa el filtro del Set (barato) y
      // recién se descarta ahí (el costo de un click de más en ese caso raro
      // es preferible a mantener un contador sincronizado a mano).
      const numContratosPendientes = new Set(contratos.map(c => normalizarEspacios(c.numContrato)));
      const { firstBtn, nextBtn, filas } = localizadoresTablaContratos(page);

      // Cota local a la cantidad de páginas, además del deadline de grupo:
      // grupoDeadlineMs (contratos.length * CONTRATO_HARD_TIMEOUT_MS) sigue
      // siendo un techo válido para el trabajo LEGÍTIMO por contrato, pero
      // es demasiado laxo para detectar el paginador atascado específico
      // que este bucle puede sufrir (nextBtn que nunca reporta disabled --
      // el mismo bug histórico documentado en findAndClickContrato, ahora
      // compartido vía avanzarPagina): para un expediente de 3,221
      // contratos esa cota son ~22 horas, mucho más que lo que esta
      // estrategia debería tardar en recorrer sus ~33 páginas reales. Con
      // ~100 filas/página confirmado en vivo, un margen generoso (2x + 10,
      // por si la tabla del sitio mezcla contratos ajenos a partida 25301
      // que no están en `contratos`) alcanza para cortar mucho antes que el
      // deadline de grupo si el paginador realmente se atoró.
      const maxPaginas = Math.ceil((totalExpediente || contratos.length) / 100) * 2 + 10;
      let paginasVisitadas = 0;

      // Fuera del loop de filas: es la misma función para las ~100 filas de
      // cada página, no hace falta recrearla en cada match.
      const abrirYCapturar = async (i) => {
        let captured = null;
        const onResponse = async (res) => {
          const m = res.url().match(detalleRe);
          if (m) {
            try { captured = { codigo: decodeURIComponent(m[1]), json: JSON.parse(await res.text()) }; }
            catch (e) { /* respuesta no-JSON, se ignora -- captured sigue null */ }
          }
        };
        page.on('response', onResponse);
        try {
          await filas.nth(i).click({ timeout: CLICK_TIMEOUT }).catch(() => {});
          const start = Date.now();
          while (!captured && Date.now() - start < RESPONSE_TIMEOUT) await page.waitForTimeout(200);
          return captured;
        } finally {
          page.off('response', onResponse);
          // closeModalIfOpen ya espera 400ms tras cerrar -- no se suma otra
          // espera fija encima (DELAY_BETWEEN_CONTRATOS_MS existía para dar
          // margen entre búsquedas que reiniciaban la tabla, no aplica aquí:
          // seguimos en la misma página, sin nada que "asentar" de más).
          await closeModalIfOpen(page);
        }
      };

      await irAPrimeraPagina(page, firstBtn);

      while (porCodigo.size > 0) {
        if (!page) break;
        if (++paginasVisitadas > maxPaginas) {
          console.error(`[pasada_unica_excedida] Expediente ${hash}: se superaron ${maxPaginas} páginas (esperadas ~${Math.ceil((totalExpediente || contratos.length) / 100)} según el tamaño del expediente) sin resolver los ${porCodigo.size} contrato(s) restantes -- posible paginador atascado, se corta acá en vez de esperar al deadline de grupo.`);
          break;
        }
        const textos = (await filas.allInnerTexts().catch(() => [])).map(normalizarEspacios);
        // Se clickea por ÍNDICE (`filas.nth(i)`), no rematcheando por texto
        // en cada iteración: si dos filas de ESTA página comparten el mismo
        // numContrato, rematchear por texto arriesgaría clickear la misma
        // fila dos veces en vez de visitar cada una -- el índice sí
        // distingue posiciones físicas distintas aunque el texto sea igual.
        // Asume que el modal (overlay `p-dialog`) no reordena las filas de
        // la tabla al abrir/cerrar dentro de la misma página -- válido para
        // un modal que se superpone sin tocar el DOM subyacente.
        for (let i = 0; i < textos.length; i++) {
          if (!page || porCodigo.size === 0) break;
          const texto = textos[i];
          if (!numContratosPendientes.has(texto)) continue;

          const captured = await withTimeout(abrirYCapturar(i), CONTRATO_HARD_TIMEOUT_MS, texto).catch(() => null);
          if (!captured || !porCodigo.has(captured.codigo)) continue; // no era ninguno de nuestros pendientes -- ver comentario de la función

          const contrato = porCodigo.get(captured.codigo);
          porCodigo.delete(captured.codigo);

          if (!captured.json.success) {
            agregarError({ hash, codigoContrato: contrato.codigoContrato, origen, error: 'respuesta_sin_exito', detalle: captured.json });
          } else if (procesarItems(contrato, captured.json.data || []) === 0) {
            agregarError({ hash, codigoContrato: contrato.codigoContrato, origen, error: 'sin_items_partida_25301_validos' });
          }
        }
        if (porCodigo.size === 0 || !page) break;
        // Progreso por página (no por contrato): una página puede resolver
        // desde 0 hasta 100 contratos según cuántos de nuestros pendientes
        // caigan ahí -- es el límite natural de esta estrategia, no un
        // conteo arbitrario como el "cada 25" del camino chico.
        console.log(`[progreso] Expediente ${hash}: ${contratos.length - porCodigo.size}/${contratos.length} contrato(s) resueltos (pasada única por página) -- ~${((Date.now() - grupoStart) / 60000).toFixed(1)} min transcurridos`);
        checkpointSiToca(ctx);
        if (!(await avanzarPagina(page, nextBtn))) break; // se recorrieron todas las páginas sin encontrar el resto
      }

      // Lo que sigue en porCodigo nunca apareció en ninguna página -- mismo
      // error que findAndClickContrato ya usa para este caso.
      for (const contrato of porCodigo.values()) {
        agregarError({ hash, codigoContrato: contrato.codigoContrato, origen, error: 'contrato_no_encontrado_en_tabla' });
      }
    }

    const procesar = (async () => {
      // totalExpediente (tamaño REAL de la tabla), no contratos.length (el
      // subconjunto pendiente esta corrida) -- un expediente gigante que el
      // checkpoint por contrato ya dejó en, digamos, 50 pendientes seguiría
      // cayendo en la búsqueda por contrato (cara, reinicia a página 1 cada
      // vez) si se comparara contra el subconjunto, justo en el caso donde
      // más rinde la pasada única: pocos pendientes dispersos en una tabla
      // igual de grande que antes.
      if (totalExpediente > UMBRAL_PASADA_UNICA) {
        await procesarPorPaginado();
        return;
      }
      for (let idx = 0; idx < contratos.length; idx++) {
        // Corte explícito, no incidental: sin este chequeo, la señal de que
        // el grupo se abandonó (page = null, asignado más abajo cuando el
        // timeout externo gana la carrera) solo llegaba como un TypeError al
        // llamar page.waitForTimeout() sobre null -- terminaba el loop igual,
        // pero por accidente, no por una condición de corte a propósito.
        if (!page) break;
        const contrato = contratos[idx];
        await page.waitForTimeout(DELAY_BETWEEN_CONTRATOS_MS);
        let res = await withTimeout(extractContrato(page, hash, contrato), CONTRATO_HARD_TIMEOUT_MS, contrato.codigoContrato)
          .catch(e => ({ error: 'excepcion', detalle: String(e) }));
        // Ver resolverAmbiguoPorPaginado: solo tiene sentido reintentar con la
        // pasada dirigida si la page sigue viva (no si el grupo ya se
        // abandonó por deadline, ver `if (!page) break` de más abajo).
        if (res.error === 'numero_contrato_ambiguo' && page) {
          res = await withTimeout(resolverAmbiguoPorPaginado(page, hash, contrato), CONTRATO_HARD_TIMEOUT_MS, contrato.codigoContrato)
            .catch(e => ({ error: 'excepcion', detalle: String(e) }));
        }
        if (res.error) {
          agregarError({ hash, codigoContrato: contrato.codigoContrato, origen, ...res });
        } else if (procesarItems(contrato, res.data) === 0) {
          agregarError({ hash, codigoContrato: contrato.codigoContrato, origen, error: 'sin_items_partida_25301_validos' });
        }

        // Progreso cada 25 contratos -- en un grupo chico (la gran mayoría del
        // CSV) esto nunca dispara, así que no ensucia el log; en un outlier de
        // miles de contratos (hasta 3,221 vistos en vivo) es la única señal de
        // avance posible hoy, ya que el checkpoint solo escribe al cerrar el
        // grupo COMPLETO -- sin esto, no había forma de distinguir "avanzando
        // lento" de "colgado" salvo midiendo CPU del proceso a mano. El ritmo y
        // ETA se calculan con el tiempo real de ESTE grupo, más preciso que
        // inferirlo de otros expedientes ya terminados (el ritmo real varía
        // mucho según qué tan rápido responde cada institución).
        if ((idx + 1) % 25 === 0) {
          const elapsedMin = (Date.now() - grupoStart) / 60000;
          const ritmo = (idx + 1) / elapsedMin;
          const etaMin = (contratos.length - (idx + 1)) / ritmo;
          const sufijoTotal = contratos.length === totalExpediente ? '' : ` (expediente completo: ${totalExpediente})`;
          console.log(`[progreso] Expediente ${hash}: ${idx + 1}/${contratos.length} contrato(s) pendientes esta corrida${sufijoTotal} (${resultados.length} registro(s), ${errores.length} error(es)) -- ~${ritmo.toFixed(1)}/min, ETA ~${etaMin.toFixed(0)} min`);
          // Mismo punto de corte que el log de arriba: checkpoint por
          // contrato DENTRO del grupo, no solo al cerrarlo -- ver comentario
          // de agregarResultado()/agregarError() más arriba.
          checkpointSiToca(ctx);
        }
      }
    })();
    // Sin este catch, un rechazo tardío del trabajo huérfano (después de que
    // el timeout de grupo ya "ganó" la carrera) sería una unhandled rejection
    // que tumba TODO el proceso -- mismo motivo que waitDetalle.catch() arriba.
    procesar.catch(() => {});
    // Promise.race no cancela al perdedor: sin clearTimeout, el timer de
    // grupoDeadlineMs queda vivo (a veces horas, para expedientes grandes)
    // aunque procesar() ya haya ganado la carrera -- Node mantiene el
    // proceso corriendo mientras haya timers pendientes, así que esto podía
    // colgar la salida del proceso mucho después de terminar todo el
    // trabajo real (candidato fuerte para los congelamientos "sin motivo"
    // observados en vivo el 2026-08-16/17).
    let deadlineTimer;
    const grupoTimedOut = await Promise.race([
      procesar.then(() => false),
      new Promise(resolve => { deadlineTimer = setTimeout(() => resolve(true), grupoDeadlineMs); }),
    ]);
    clearTimeout(deadlineTimer);
    if (grupoTimedOut) {
      console.error(`[grupo_timeout] Expediente ${hash} excedió ${grupoDeadlineMs}ms con ${fmtContratos(contratos.length, totalExpediente)} -- se fuerza cierre de page y se abandona el resto del expediente para esta corrida.`);
      await page.close().catch(() => {});
      page = null;
      // Ventana corta de gracia antes de dar por perdido el trabajo huérfano:
      // cerrar la page suele hacer que cualquier operación de Playwright en
      // vuelo contra ella (el contrato que se estaba procesando justo cuando
      // venció el deadline) rechace casi de inmediato ("Target closed" y
      // similares). Con agregarResultado()/agregarError() empujando directo a
      // ctx.allResultados/ctx.allErrores (checkpoint por contrato, ver arriba)
      // una resolución tardía YA NO se pierde aunque llegue después de esta
      // ventana -- el array compartido sigue vivo y se sigue escribiendo a
      // disco más adelante, a diferencia de antes (cuando worker() copiaba
      // los arrays locales por valor una sola vez al retornar). Y si de
      // todos modos llega DESPUÉS de esta ventana (una vez que el error
      // sintético 'grupo_timeout' de abajo ya se generó), agregarResultado()/
      // agregarError() lo descartan en vez de dejar un contrato duplicado
      // (ver descartarTimeoutSintetico() arriba). Esta espera de 2s ya no es
      // la única defensa contra duplicados/pérdida -- solo reduce cuán
      // seguido hace falta que la de-dup entre en acción.
      await Promise.race([procesar, new Promise(r => setTimeout(r, 2000))]).catch(() => {});
      // Cuentan como error real (contra MAX_INTENTOS), no solo "no llegamos"
      // -- ver comentario arriba de por qué llegar acá ya es señal de algo
      // roto. Solo los que ni tienen resultado ni ya generaron su propio
      // error dentro del loop (el contrato que estaba en vuelo justo cuando
      // cerró la page sí puede haber quedado con un error 'excepcion' propio
      // -- no se le pisa ni se le duplica).
      const yaResueltos = new Set(resultados.map(r => r.codigo_contrato));
      const yaConError = new Set(errores.map(e => e.codigoContrato));
      for (const contrato of contratos) {
        if (yaResueltos.has(contrato.codigoContrato) || yaConError.has(contrato.codigoContrato)) continue;
        agregarError({ hash, codigoContrato: contrato.codigoContrato, origen, error: 'grupo_timeout', detalle: `expediente abandonado por deadline de grupo (${grupoDeadlineMs}ms)` });
      }
      checkpointSiToca(ctx);
    }
  }

  if (page) await page.close().catch(() => {});

  ctx.done++;
  console.log(`[${ctx.done}/${ctx.total}] Expediente ${hash} (${fmtContratos(contratos.length, totalExpediente)}) -> ${resultados.length} registro(s), ${errores.length} error(es)`);

  // Sin valor de retorno: resultados, errores y stats ya se empujaron en
  // vivo a ctx.allResultados/ctx.allErrores/ctx.stats vía
  // agregarResultado()/agregarError()/buildRegistro() arriba -- worker() no
  // necesita (ni debe) volver a fundir nada.
}

function guardarCheckpoint(outPath, errPath, rawPath, resultados, errores, rawItems) {
  // Sin indentar: esto reescribe el dataset COMPLETO (decenas de miles de
  // registros) en cada checkpoint -- con el checkpoint por contrato, eso
  // puede disparar varias veces por expediente grande, no solo entre
  // expedientes. Indentado no aporta nada (nadie revisa este JSON línea por
  // línea a este tamaño) y cuesta stringify + I/O notablemente más caro en
  // el hot path.
  fs.writeFileSync(outPath, JSON.stringify(resultados), 'utf8');
  fs.writeFileSync(errPath, JSON.stringify(errores), 'utf8');
  fs.writeFileSync(rawPath, JSON.stringify(rawItems), 'utf8');
}

// Por tiempo, no por cantidad de expedientes/contratos -- reescribir el JSON
// completo en cada llamada tiene costo O(n), así que dispararlo demasiado
// seguido en ráfagas (muchos expedientes chicos, o el loop de checkpoint por
// contrato dentro de un grupo grande) degeneraría en un patrón O(n²) sobre
// toda la corrida. `ctx` centraliza el estado compartido entre workers
// (`ultimoCheckpoint`) y los arrays/paths de salida, así que cualquier punto
// de la corrida -- entre expedientes o DENTRO de uno grande -- puede pedir
// "checkpointeá si toca" sin duplicar la lógica del intervalo.
function checkpointSiToca(ctx) {
  const ahora = Date.now();
  if (ahora - ctx.ultimoCheckpoint < CHECKPOINT_MIN_INTERVALO_MS) return;
  ctx.ultimoCheckpoint = ahora;
  guardarCheckpoint(ctx.outPath, ctx.errPath, ctx.rawPath, ctx.allResultados, ctx.allErrores, ctx.allRawItems);
}

async function worker(queue, browser, compendio, cucopMap, ctx, origen) {
  while (queue.length) {
    const [hash, contratos, totalExpediente] = queue.shift();
    // Log de arranque (no solo de fin): sin esto, un expediente que tarda
    // mucho (uno grande, o uno realmente colgado) es indistinguible desde
    // afuera -- no hay forma de saber qué está en vuelo durante un silencio
    // largo en el log, solo cuál fue el último que terminó.
    console.log(`[inicio] Expediente ${hash} (${fmtContratos(contratos.length, totalExpediente)}) a las ${new Date().toISOString()}`);
    await processExpedienteGroup(browser, hash, contratos, compendio, cucopMap, ctx, origen, totalExpediente);

    // Cola de seguridad: la mayoría del progreso de este grupo ya se
    // checkpointeó EN VIVO dentro de processExpedienteGroup (ver
    // checkpointSiToca ahí) -- esto solo cubre el resto que quedó pendiente
    // desde el último checkpoint interno hasta que el grupo cerró.
    checkpointSiToca(ctx);

    await new Promise(r => setTimeout(r, DELAY_BETWEEN_EXPEDIENTES_MS));
  }
}

function estaVivo(pid) {
  try {
    process.kill(pid, 0);
  } catch (e) {
    return false;
  }
  // process.kill(pid, 0) solo confirma que ALGÚN proceso vive con ese PID --
  // Windows recicla PIDs, así que un lock viejo podría "encontrar vivo" un
  // proceso no relacionado que por casualidad heredó el mismo número, y
  // bloquear todo relanzamiento futuro para este origen hasta borrar el
  // lock a mano. Se confirma además que ese PID es un node.exe (el runtime
  // de extract.js) antes de darlo por válido -- no es una prueba perfecta,
  // pero reduce mucho la ventana de falso positivo.
  try {
    const salida = execSync(`tasklist /FI "PID eq ${pid}" /FO CSV /NH`, { encoding: 'utf8' });
    return /node\.exe/i.test(salida);
  } catch (e) {
    return false;
  }
}

// Evita dos corridas concurrentes sobre el mismo `origen` escribiendo al
// mismo docs/data.json -- pasó en vivo el 2026-08-11: "killed" del harness de
// background tasks no mata de verdad al proceso hijo en Windows, y se
// relanzó sin verificar primero, terminando con 3 corridas de extract.js
// pisándose el progreso entre sí (last-write-wins en cada checkpoint). Con
// este lock, un relanzamiento mientras la corrida real sigue viva falla
// rápido y no toca nada, en vez de competir por escribir el mismo archivo.
function adquirirLock(lockPath) {
  if (fs.existsSync(lockPath)) {
    const pidPrevio = parseInt(fs.readFileSync(lockPath, 'utf8').trim(), 10);
    if (pidPrevio && estaVivo(pidPrevio)) {
      console.error(`ERROR: ya hay una corrida en curso para este origen (PID ${pidPrevio}, lock: ${lockPath}). No se relanza para evitar pisar su progreso.`);
      process.exit(1);
    }
    console.log(`Lock huérfano encontrado (PID ${pidPrevio}, ya no existe o no es un proceso node) -- se descarta y se continúa.`);
    fs.unlinkSync(lockPath);
  }
  // 'wx' (exclusive create) en vez de un writeFileSync normal: existsSync +
  // writeFileSync por separado deja una ventana de carrera real entre el
  // chequeo y la escritura -- dos procesos que arrancan con milisegundos de
  // diferencia podrían pasar ambos el existsSync antes de que cualquiera
  // escriba, y los dos terminarían creyendo que ganaron el lock. Con 'wx',
  // la escritura misma falla con EEXIST si alguien ganó la carrera primero.
  try {
    fs.writeFileSync(lockPath, String(process.pid), { flag: 'wx' });
  } catch (e) {
    if (e.code === 'EEXIST') {
      console.error(`ERROR: otra corrida ganó el lock justo antes (${lockPath}). No se relanza para evitar pisar su progreso.`);
      process.exit(1);
    }
    throw e;
  }
}

function liberarLock(lockPath) {
  try { fs.unlinkSync(lockPath); } catch (e) { /* ya no existe, no pasa nada */ }
}

function cargarEstadoPrevio(outPath, errPath, rawPath) {
  const resultados = fs.existsSync(outPath) ? JSON.parse(fs.readFileSync(outPath, 'utf8')) : [];
  const errores = fs.existsSync(errPath) ? JSON.parse(fs.readFileSync(errPath, 'utf8')) : [];
  const rawItems = fs.existsSync(rawPath) ? JSON.parse(fs.readFileSync(rawPath, 'utf8')) : [];
  return { resultados, errores, rawItems };
}

// Decide qué expedientes hay que (re)procesar en esta corrida, y con qué
// contratos exactos. Por cada contrato objetivo puede estar en uno de tres
// estados:
//   - "resuelto": ya tiene al menos un registro exitoso en resultados previos.
//   - "fallo permanente": no tiene resultados y ya agotó MAX_INTENTOS errores.
//   - "pendiente": todo lo demás -- nunca se intentó (corrida nueva o
//     contrato recién aparecido en el CSV), o falló pero todavía puede
//     reintentarse.
// Un expediente se salta solo si TODOS sus contratos están resueltos o en
// fallo permanente. Si al menos uno está pendiente, `pendientes` guarda SOLO
// el subconjunto pendiente de ese expediente (no el grupo completo) -- hay
// que recargar la página igual (no hay forma de evitar eso), pero
// processExpedienteGroup ya no vuelve a scrapear contratos que ya tuvieron
// éxito o que ya agotaron sus reintentos. Antes reprocesar el expediente
// COMPLETO por un solo contrato pendiente tiraba a la basura el progreso ya
// bueno de todo el resto del grupo en cada corte -- con expedientes de hasta
// 2,457 contratos (corrida de 2025), eso significaba re-scrapear horas de
// trabajo ya hecho cada vez que la ventana de proceso cortaba cerca del
// final. Ver también el ajuste correspondiente en main() (ya no descarta
// resultados previos de contratos resueltos dentro de un grupo pendiente).
function calcularPendientes(groups, resultadosPrevios, erroresPrevios) {
  const codigosConResultado = new Set(resultadosPrevios.map(r => r.codigo_contrato));
  // `erroresPrevios` trae UNA entrada por contrato (ver agregarError), con su
  // propio contador `intentos` -- no hace falta contar duplicados acá, cada
  // corrida ya consolidó lo suyo en esa única entrada.
  const intentosPorContrato = new Map();
  for (const e of erroresPrevios) {
    intentosPorContrato.set(`${e.hash}|${e.codigoContrato}`, e.intentos || 1);
  }

  const pendientes = new Map();
  let saltados = 0, fallosPermanentes = 0;
  for (const [hash, contratos] of groups) {
    const contratosPendientes = contratos.filter(c => {
      if (codigosConResultado.has(c.codigoContrato)) return false;
      const intentos = intentosPorContrato.get(`${hash}|${c.codigoContrato}`) || 0;
      return intentos < MAX_INTENTOS;
    });

    if (contratosPendientes.length > 0) {
      pendientes.set(hash, contratosPendientes);
    } else if (contratos.every(c => codigosConResultado.has(c.codigoContrato))) {
      saltados++;
    } else {
      fallosPermanentes++;
    }
  }
  return { pendientes, saltados, fallosPermanentes };
}

function construirReporteCalidad(resultados, stats) {
  // valor_maximo/cantidad_maxima representan el "número único" de referencia
  // de cada registro (equivalen al antiguo `valor`/`cantidad` -- iguales a
  // valor_minimo/cantidad_minima cuando no hay rango genuino, ver
  // buildRegistro), así que sirven igual para ordenar outliers y detectar ceros.
  const topValores = [...resultados].sort((a, b) => (b.valor_maximo || 0) - (a.valor_maximo || 0)).slice(0, 20)
    .map(r => ({ codigo_contrato: r.codigo_contrato, producto: r.producto.slice(0, 80), precio_unitario: r.precio_unitario, cantidad_maxima: r.cantidad_maxima, valor_maximo: r.valor_maximo }));
  const enCero = resultados.filter(r => (r.precio_unitario ?? 0) <= 0 || (r.cantidad_maxima ?? 0) <= 0)
    .map(r => ({ codigo_contrato: r.codigo_contrato, producto: r.producto.slice(0, 80), precio_unitario: r.precio_unitario, cantidad_maxima: r.cantidad_maxima }));

  return {
    generado: new Date().toISOString(),
    total_registros: resultados.length,
    top_20_valores_mas_altos: topValores,
    registros_con_precio_o_cantidad_en_cero_o_negativo: enCero,
    precios_corregidos_por_inconsistencia_con_subtotal: stats.preciosCorregidos,
    cantidades_derivadas_de_subtotal_entre_precio_unitario: stats.cantidadesDerivadas,
  };
}

async function main() {
  const args = parseArgs(process.argv);
  if (!args.input) {
    console.error('Uso: node scripts/extract.js <archivo.csv> [--limit N] [--out archivo.json] [--concurrency N] [--reset]');
    process.exit(1);
  }
  console.log(`Origen de esta corrida: "${args.origen}"`);

  const lockPath = path.join(__dirname, '..', 'data', `.extract-${args.origen}.lock`);
  adquirirLock(lockPath);
  process.on('exit', () => liberarLock(lockPath));

  console.log(`Cargando compendio de medicamentos...`);
  const compendio = loadCompendio();
  console.log(`Compendio cargado: ${Object.keys(compendio).length} claves.`);

  console.log(`Cargando mapa CUCoP+ -> clave CSG (data/cucop_medicamentos.json)...`);
  const cucopMap = loadCucopMap();
  console.log(`Mapa cargado: ${Object.keys(cucopMap).length} claves CUCoP+.`);

  console.log(`Leyendo y filtrando CSV: ${args.input}`);
  const filtered = loadFilteredRows(args.input, args.limit);
  console.log(`Filas filtradas (partida 25301, con datos completos): ${filtered.length}`);

  const groups = groupByHash(filtered);
  console.log(`Agrupadas en ${groups.size} expedientes.`);

  // Fijos (no derivados de args.out): a diferencia de los registros, que se
  // publican un archivo por año (ver parseArgs), el historial de
  // errores/items crudos sigue combinado en un solo archivo entre todos los
  // orígenes -- son chicos (no se acercan al límite de 100 MB de GitHub) y
  // calcularPendientes() necesita ver los intentos de TODOS los orígenes
  // para contar bien MAX_INTENTOS entre corridas.
  const docsDir = path.dirname(args.out);
  const errPath = path.join(docsDir, 'data.errores.json');
  const rawPath = path.join(docsDir, 'data.raw.json');
  let { resultados: resultadosPrevios, errores: erroresPrevios, rawItems: rawItemsPrevios } = cargarEstadoPrevio(args.out, errPath, rawPath);
  console.log(`Estado previo: ${resultadosPrevios.length} registros, ${erroresPrevios.length} errores acumulados, ${rawItemsPrevios.length} items crudos.`);

  if (args.reset) {
    const antesR = resultadosPrevios.length;
    const antesE = erroresPrevios.length;
    const antesRaw = rawItemsPrevios.length;
    // Solo toca registros/errores marcados con ESTE origen -- cualquier otro
    // origen (u origen ausente, dato legado) queda intacto sin importar su
    // codigo_contrato. Ver nota en parseArgs() sobre por qué no se usa el
    // prefijo de codigo_contrato para esto.
    resultadosPrevios = resultadosPrevios.filter(r => r.origen !== args.origen);
    erroresPrevios = erroresPrevios.filter(e => e.origen !== args.origen);
    rawItemsPrevios = rawItemsPrevios.filter(r => r.origen !== args.origen);
    console.log(`--reset: descartados ${antesR - resultadosPrevios.length} registros, ${antesE - erroresPrevios.length} errores y ${antesRaw - rawItemsPrevios.length} items crudos con origen "${args.origen}" (otros orígenes intactos).`);
  }

  const { pendientes, saltados, fallosPermanentes } = calcularPendientes(groups, resultadosPrevios, erroresPrevios);
  console.log(`${saltados} expedientes ya resueltos (se saltan), ${fallosPermanentes} con fallo permanente (>= ${MAX_INTENTOS} intentos, se saltan), ${pendientes.size} pendientes de procesar. Concurrencia: ${args.concurrency}`);

  // Se descartan los RESULTADOS previos SOLO de los contratos que se van a
  // reprocesar esta corrida (pendientes.values() ya viene filtrado a ese
  // subconjunto, ver calcularPendientes), no de todo el expediente al que
  // pertenecen -- así un contrato ya resuelto dentro de un grupo que todavía
  // tiene otros pendientes conserva su resultado en vez de perderlo y tener
  // que ser re-scrapeado de nuevo esta corrida.
  const codigosPendientes = new Set();
  for (const contratos of pendientes.values()) for (const c of contratos) codigosPendientes.add(c.codigoContrato);
  const resultados = resultadosPrevios.filter(r => !codigosPendientes.has(r.codigo_contrato));
  // Los ERRORES previos, en cambio, NO se descartan aquí -- deben persistir
  // completos entre corridas para que calcularPendientes() pueda contar
  // intentos acumulados y aplicar MAX_INTENTOS. Antes esta línea también
  // filtraba erroresPrevios igual que resultadosPrevios: cualquier contrato
  // que siguiera pendiente perdía TODO su historial de error al arrancar la
  // siguiente corrida, así que intentosPorContrato nunca pasaba de 1 y un
  // contrato que fallaba siempre de la misma forma (ver
  // sin_items_partida_25301_validos) quedaba reintentándose para siempre, sin
  // llegar nunca a "fallo permanente" -- confirmado en vivo el 2026-08-21
  // corriendo el scrape completo de 2026. descartarTimeoutSintetico() arriba
  // ya cubre el único caso real de duplicado que hacía falta evitar (el
  // marcador sintético de 'grupo_timeout' superado por el resultado real),
  // así que no hace falta ningún descarte adicional acá.
  const errores = erroresPrevios;
  const rawItems = rawItemsPrevios.filter(r => !codigosPendientes.has(r.codigo_contrato));

  const stats = { preciosCorregidos: [], cantidadesDerivadas: [] };
  const browser = await chromium.launch({ headless: true });

  // Expedientes chicos primero: la mayoría tiene pocos contratos y termina
  // (y checkpointea) en segundos, pero unos pocos outliers tienen cientos o
  // miles (visto en vivo: hasta 2,457 en la corrida de 2025) -- con
  // CONTRATO_HARD_TIMEOUT_MS por contrato, uno de esos solo puede tardar
  // horas. Sin este orden, un worker que agarra un outlier temprano queda
  // "atado" a él durante toda la corrida, y si la ventana de proceso se corta
  // antes de que termine, se pierde y se reintenta desde cero -- procesar los
  // chicos primero maximiza el progreso que sobrevive a un corte, y deja los
  // outliers para el final (donde tienen más chance de completarse solos si
  // ya no compiten por los 4 workers).
  // Tamaño REAL de cada expediente (antes de recortar a pendientes), en la
  // propia tupla de la cola -- solo para mostrarlo en los logs. Desde el fix
  // de checkpoint quirúrgico, `contratos` que le llega a
  // processExpedienteGroup es el subconjunto pendiente, no el grupo
  // completo; sin esto, un expediente de miles de contratos que ya tiene la
  // mayoría resuelto de una corrida anterior aparece en el log como "3
  // contrato(s)", indistinguible de uno genuinamente chico -- dificulta
  // diagnosticar cuáles son los outliers reales mirando el log en vivo.
  const queue = [...pendientes.entries()]
    .map(([hash, contratosPendientes]) => [hash, contratosPendientes, groups.get(hash).length])
    .sort((a, b) => a[1].length - b[1].length);
  // `ctx` centraliza el estado compartido entre workers concurrentes -- los
  // arrays/paths de salida y `stats` para que checkpointSiToca()/
  // buildRegistro() puedan usarse desde cualquier punto de la corrida (entre
  // expedientes o DENTRO de uno grande, ver processExpedienteGroup) sin
  // pasar varios parámetros más por cada función de la cadena.
  const ctx = { done: 0, total: queue.length, ultimoCheckpoint: 0, allResultados: resultados, allErrores: errores, allRawItems: rawItems, outPath: args.out, errPath, rawPath, stats };

  const workers = Array.from({ length: args.concurrency }, () =>
    worker(queue, browser, compendio, cucopMap, ctx, args.origen)
  );
  await Promise.all(workers);

  await browser.close();

  guardarCheckpoint(args.out, errPath, rawPath, resultados, errores, rawItems);

  // calidad.json y data.xlsx cubren el dataset COMPLETO (todos los orígenes),
  // no solo el de esta corrida -- se recarga todo desde disco (incluye lo que
  // guardarCheckpoint acaba de escribir arriba) en vez de necesitar cargar
  // otros orígenes en memoria durante toda la corrida. `stats` (precios
  // corregidos/cantidades derivadas) sigue siendo solo de esta corrida --
  // esas cifras son sobre trabajo hecho ahora, no sobre el dataset entero.
  const datasetCompleto = cargarDatasetCompleto(docsDir);
  guardarManifest(docsDir);

  const reporte = construirReporteCalidad(datasetCompleto, stats);
  const reportePath = path.join(docsDir, 'data.calidad.json');
  fs.writeFileSync(reportePath, JSON.stringify(reporte, null, 2), 'utf8');

  const excelPath = path.join(docsDir, 'data.xlsx');
  await guardarExcel(excelPath, datasetCompleto);

  console.log('--- Resumen ---');
  console.log(`Expedientes procesados en esta corrida: ${ctx.total}`);
  console.log(`Registros de origen "${args.origen}": ${resultados.length}`);
  console.log(`Registros totales (todos los orígenes): ${datasetCompleto.length}`);
  console.log(`Errores totales acumulados: ${errores.length}`);
  console.log(`Precios corregidos por inconsistencia esta corrida: ${stats.preciosCorregidos.length}`);
  console.log(`Cantidades derivadas de subtotal/precio_unitario esta corrida: ${stats.cantidadesDerivadas.length}`);
  console.log(`Salida: ${args.out}`);
  console.log(`Excel: ${excelPath}`);
  console.log(`Errores detallados: ${errPath}`);
  console.log(`Items crudos (${rawItems.length}): ${rawPath}`);
  console.log(`Reporte de calidad: ${reportePath}`);
}

main().catch(err => {
  console.error('ERROR FATAL:', err);
  process.exit(1);
});
