// Pipeline principal: filtra el CSV masivo de Compras MX por medicamentos (partida 25301),
// extrae precio unitario por contrato via Playwright, cruza contra el Compendio Nacional
// de Medicamentos, y guarda el resultado consolidado.
//
// Uso:
//   node scripts/extract.js data/raw/contratos_2026.csv [--limit N] [--out data/precios.json]

const fs = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');
const { chromium } = require('playwright');
const { CLAVE_RE, guardarExcel } = require('./lib/dataset');

// La institución a veces captura "Descripción detallada" del requerimiento de
// forma degenerada: como mera referencia a su propia partida, sin texto de
// producto ("CONFORME A PARTIDA 3 DE LA CONVOCATORIA"). Verificado en vivo
// contra el sitio (contrato C-2026-00073692): el desglose real del producto
// no está en ningún otro campo de ese contrato, solo en la ficha CUCoP+ de
// la convocatoria (ver PRODUCTO_SIN_ESPACIOS_RE abajo y buildRegistro).
const CONFORME_PARTIDA_RE = /^CONFORME\s+A\s+PARTIDA\s+\d+\s+DE\s+LA\s+CONVOCATORIA$/i;
// Otras veces sí trae el producto, pero sin ningún espacio entre palabras --
// visto en vivo en procedimiento IA-13-312-013000999-T-327-2026: los 174
// ítems de "Descripción detallada" vienen así, de origen, mientras que
// "Descripción CUCoP+" (misma página) sí viene bien formada. Umbral de 25
// caracteres para evitar falsos positivos en descripciones cortas
// legítimamente compactas.
const PRODUCTO_SIN_ESPACIOS_RE = (s) => s.length > 25 && !s.includes(' ');
// Variante del caso de arriba: la institución SÍ describe el producto
// completo y bien formado, pero le pega la misma coletilla administrativa al
// final ("...ENVASE CON 10 TABLETAS.CONFORME A PARTIDA 204 DE LA
// CONVOCATORIA"), con o sin punto/espacio de separación, a veces sin el
// número de partida. A diferencia de CONFORME_PARTIDA_RE (arriba), aquí NO
// se reemplaza todo el producto por la ficha CUCoP+ -- la descripción propia
// ya es buena, solo se recorta la coletilla. Verificado 2026-08-10: 612 casos
// reales, ninguno coincide también con CONFORME_PARTIDA_RE (poblaciones
// disjuntas -- ver Limitaciones.md §7).
const CONFORME_PARTIDA_SUFIJO_RE = /\s*[.\-•]?\s*CONFORME\s+A\s+PARTIDA\s*\d*\s*DE\s+LA\s+CONVOCATORIA\.?\s*$/i;
// Numeración de partida, viñetas o claves con separador distinto al esperado
// que algunas instituciones anteponen a "Descripción detallada", ajenas al
// nombre del medicamento -- ej. "13040096 UPADACITINIB...", "2. 1
// TEOFILINA...", "-OLANZAPINA...", "•\tASPIRINA...", "10.   010.000.2739.00
// - DIETA...". Exigir que el prefijo termine justo antes de una letra evita
// comerse dígitos que sí son parte del texto (ej. "5% DIOXIDO DE
// CARBONO..." -- el "5" no se toca porque le sigue "%", no una letra). La
// guarda negativa evita truncar a la mitad un sufijo de clave alfanumérico
// (ej. "...1757.P0 MELFALAN..." se deja intacto en vez de dejar "P0
// MELFALAN...").
const LEADING_JUNK_RE = /^[\d\s.\-•"“]+(?=[A-Za-zÀ-ÿ])(?![A-Za-zÀ-ÿ]\d\s)/;
const HASH_RE = /detalle\/([a-f0-9]+)\/procedimiento/i;
const NAV_TIMEOUT = 30000;
const CLICK_TIMEOUT = 15000;
const RESPONSE_TIMEOUT = 10000;
const DELAY_BETWEEN_CONTRATOS_MS = 400;
const DELAY_BETWEEN_EXPEDIENTES_MS = 800;
const CONTRATO_HARD_TIMEOUT_MS = 25000; // nunca dejar que un solo contrato cuelgue el pipeline (sube un poco vs. antes: ahora la búsqueda es exhaustiva y puede recorrer más páginas)
const MAX_INTENTOS = 3; // tras esto, un expediente con error se marca como fallo permanente y ya no se reintenta solo
const CHECKPOINT_CADA_N_EXPEDIENTES = 10; // escritura incremental a disco, para no perder progreso si el proceso se corta

function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(`hard_timeout: ${label}`)), ms)),
  ]);
}

function parseArgs(argv) {
  const args = { input: null, limit: Infinity, out: path.join(__dirname, '..', 'docs', 'data.json'), concurrency: 1 };
  const rest = argv.slice(2);
  args.input = rest.find(a => !a.startsWith('--'));
  const limitIdx = rest.indexOf('--limit');
  if (limitIdx !== -1) args.limit = parseInt(rest[limitIdx + 1], 10);
  const outIdx = rest.indexOf('--out');
  if (outIdx !== -1) args.out = rest[outIdx + 1];
  const concIdx = rest.indexOf('--concurrency');
  if (concIdx !== -1) args.concurrency = parseInt(rest[concIdx + 1], 10);
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

async function findAndClickContrato(page, numContrato) {
  // Los botones de paginación se acotan al que sigue al encabezado
  // "DATOS RELEVANTES DEL CONTRATO" en el DOM, porque la página puede tener
  // varias tablas paginadas (anexos, requerimientos, contratos) con la misma clase.
  const scope = page.locator('text=DATOS RELEVANTES DEL CONTRATO');
  const firstBtn = scope.locator('xpath=following::*[contains(@class,"p-paginator-first")][1]');
  const nextBtn = scope.locator('xpath=following::*[contains(@class,"p-paginator-next")][1]');

  // Reiniciar siempre a la primera página: como esta función se llama una vez
  // por cada contrato objetivo dentro del mismo expediente, sin este reset la
  // búsqueda del segundo contrato en adelante arrancaría desde donde quedó la
  // paginación tras buscar el anterior, en vez de desde el principio.
  const firstDisabled = await firstBtn.evaluate(el => el.classList.contains('p-disabled')).catch(() => true);
  if (!firstDisabled) {
    await firstBtn.click({ timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(600);
  }

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
    const disabled = await nextBtn.evaluate(el => el.classList.contains('p-disabled')).catch(() => true);
    if (disabled) return false; // se recorrieron todas las páginas, no está en la tabla
    await nextBtn.click({ timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(600);
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

async function extractContrato(page, hash, contrato) {
  let captured = null;
  const onResponse = async (res) => {
    if (res.url().includes(`/whitney/sitiopublico/detallepartidas/${hash}/${contrato.codigoContrato}`)) {
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
    if (!captured.success) return { error: 'respuesta_sin_exito', detalle: captured };

    return { data: captured.data || [] };
  } finally {
    page.off('response', onResponse);
    await closeModalIfOpen(page);
  }
}

function buildRegistro(contrato, item, compendio, cucopMap, stats) {
  // Todo lo que llega aquí ya viene de un contrato con Partida específica 25301
  // (medicamentos), así que se conserva aunque no traiga la clave del compendio
  // embebida en la descripción -- solo algunas instituciones (grandes/federales)
  // la citan; el resto no. Cuando SÍ hay clave y existe en el compendio, se
  // enriquece con grupo_terapéutico; si no, quedan en null (baja confianza pero
  // se conserva la fila).
  const m = String(item.descripcion || '').match(CLAVE_RE);
  let producto = m ? m[5].trim() : String(item.descripcion || '').trim();

  // Cuando la descripción de la institución es degenerada (ver
  // CONFORME_PARTIDA_RE / PRODUCTO_SIN_ESPACIOS_RE arriba), se reemplaza por
  // la ficha oficial del catálogo CUCoP+ -- ya cargada en cucopMap para
  // resolver `clave`, siempre bien formada. No todos los cve_cucop están en
  // el catálogo local (puede ser más nuevo que el snapshot descargado); en
  // esos casos queda el texto degenerado tal cual, no hay mejor fuente.
  if (CONFORME_PARTIDA_RE.test(producto) || PRODUCTO_SIN_ESPACIOS_RE(producto)) {
    const viaCucop = cucopMap[item.cve_cucop];
    if (viaCucop) producto = viaCucop.descripcion;
  }

  // Coletilla administrativa pegada al final de una descripción por lo demás
  // buena (ver CONFORME_PARTIDA_SUFIJO_RE arriba) -- se recorta, no se
  // reemplaza todo el producto como en el caso de arriba. Corre después del
  // bloque anterior: los casos puros (producto == solo la frase) ya quedaron
  // reemplazados por su ficha CUCoP+, que nunca contiene esta frase, así que
  // este paso no les afecta.
  producto = producto.replace(CONFORME_PARTIDA_SUFIJO_RE, '').trim();

  producto = producto.replace(LEADING_JUNK_RE, '').trim();

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
  };
}

async function processExpedienteGroup(browser, hash, contratos, compendio, cucopMap, stats, ctx) {
  const url = `https://comprasmx.buengobierno.gob.mx/sitiopublico/#/sitiopublico/detalle/${hash}/procedimiento`;
  const resultados = [];
  const errores = [];

  // Pestaña nueva por expediente: reutilizar una sola pestaña entre navegaciones
  // de rutas con hash de esta SPA (Angular) resultó no confiable en pruebas.
  const page = await browser.newPage();

  let loaded = false;
  for (let intento = 0; intento < 2 && !loaded; intento++) {
    try {
      const waitDetalle = page.waitForResponse(
        res => res.url().includes(`/whitney/sitiopublico/expedientes/${hash}?id_proceso=procedimiento`) && res.status() === 200,
        { timeout: NAV_TIMEOUT }
      );
      await page.goto(url, { waitUntil: 'load', timeout: NAV_TIMEOUT });
      await waitDetalle;
      await page.waitForTimeout(800); // margen para que Angular renderice la tabla tras la respuesta
      loaded = true;
    } catch (e) {
      if (intento === 1) errores.push({ hash, error: 'navegacion_fallida', detalle: String(e) });
    }
  }

  if (loaded) {
    let contratoIdx = 0;
    for (const contrato of contratos) {
      contratoIdx++;
      await page.waitForTimeout(DELAY_BETWEEN_CONTRATOS_MS);
      const res = await withTimeout(extractContrato(page, hash, contrato), CONTRATO_HARD_TIMEOUT_MS, contrato.codigoContrato)
        .catch(e => ({ error: 'excepcion', detalle: String(e) }));
      if (res.error) {
        errores.push({ hash, codigoContrato: contrato.codigoContrato, ...res });
        continue;
      }
      for (const item of res.data) {
        // El filtro de la partida 25301 en loadFilteredRows() opera a nivel
        // de CONTRATO (una fila del CSV masivo), no de ítem: un contrato con
        // al menos una línea de partida 25301 pasa el filtro, pero
        // detallepartidas() devuelve TODAS las líneas del contrato, incluidas
        // las de otras partidas (radiofármacos 25401, material de curación
        // 25501, etc.) si la compra fue mixta. Sin este segundo filtro por
        // ítem, esas líneas se colaban al dataset marcadas como MEDICAMENTO
        // (~32% del dataset en la corrida de 2026-08-09, detectado por cve_cucop
        // fuera de 25301-*).
        if (!item.cve_cucop || !String(item.cve_cucop).startsWith('25301')) continue;
        const registro = buildRegistro(contrato, item, compendio, cucopMap, stats);
        if (registro) resultados.push(registro);
      }
    }
  }

  await page.close();

  ctx.done++;
  console.log(`[${ctx.done}/${ctx.total}] Expediente ${hash} (${contratos.length} contrato(s)) -> ${resultados.length} registro(s), ${errores.length} error(es)`);

  return { resultados, errores };
}

function guardarCheckpoint(outPath, errPath, resultados, errores) {
  fs.writeFileSync(outPath, JSON.stringify(resultados, null, 2), 'utf8');
  fs.writeFileSync(errPath, JSON.stringify(errores, null, 2), 'utf8');
}

async function worker(queue, browser, compendio, cucopMap, stats, ctx, allResultados, allErrores, outPath, errPath) {
  while (queue.length) {
    const [hash, contratos] = queue.shift();
    const { resultados, errores } = await processExpedienteGroup(browser, hash, contratos, compendio, cucopMap, stats, ctx);
    allResultados.push(...resultados);
    allErrores.push(...errores);

    // Checkpoint incremental: si el proceso se corta (timeout, caída, cierre
    // manual), no se pierde todo el progreso -- la siguiente corrida retoma
    // desde el último checkpoint escrito, no desde cero.
    if (ctx.done % CHECKPOINT_CADA_N_EXPEDIENTES === 0) {
      guardarCheckpoint(outPath, errPath, allResultados, allErrores);
    }

    await new Promise(r => setTimeout(r, DELAY_BETWEEN_EXPEDIENTES_MS));
  }
}

function cargarEstadoPrevio(outPath, errPath) {
  const resultados = fs.existsSync(outPath) ? JSON.parse(fs.readFileSync(outPath, 'utf8')) : [];
  const errores = fs.existsSync(errPath) ? JSON.parse(fs.readFileSync(errPath, 'utf8')) : [];
  return { resultados, errores };
}

// Decide qué expedientes hay que (re)procesar en esta corrida. Por cada
// contrato objetivo puede estar en uno de tres estados:
//   - "resuelto": ya tiene al menos un registro exitoso en resultados previos.
//   - "fallo permanente": no tiene resultados y ya agotó MAX_INTENTOS errores.
//   - "pendiente": todo lo demás -- nunca se intentó (corrida nueva o
//     contrato recién aparecido en el CSV), o falló pero todavía puede
//     reintentarse.
// Un expediente se salta solo si TODOS sus contratos están resueltos o en
// fallo permanente; si al menos uno está pendiente, se reprocesa el
// expediente COMPLETO (más simple que trackear parcialidad dentro de él, y
// de todas formas hay que recargar la página entera).
function calcularPendientes(groups, resultadosPrevios, erroresPrevios) {
  const codigosConResultado = new Set(resultadosPrevios.map(r => r.codigo_contrato));
  const intentosPorContrato = new Map();
  for (const e of erroresPrevios) {
    const key = `${e.hash}|${e.codigoContrato}`;
    intentosPorContrato.set(key, (intentosPorContrato.get(key) || 0) + 1);
  }

  const pendientes = new Map();
  let saltados = 0, fallosPermanentes = 0;
  for (const [hash, contratos] of groups) {
    const estados = contratos.map(c => {
      if (codigosConResultado.has(c.codigoContrato)) return 'resuelto';
      const intentos = intentosPorContrato.get(`${hash}|${c.codigoContrato}`) || 0;
      if (intentos >= MAX_INTENTOS) return 'fallo_permanente';
      return 'pendiente'; // nunca intentado, o falló pero aún reintentable
    });

    if (estados.some(e => e === 'pendiente')) {
      pendientes.set(hash, contratos);
    } else if (estados.every(e => e === 'resuelto')) {
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
    console.error('Uso: node scripts/extract.js <archivo.csv> [--limit N] [--out archivo.json] [--concurrency N]');
    process.exit(1);
  }

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

  const errPath = args.out.replace(/\.json$/, '.errores.json');
  const { resultados: resultadosPrevios, errores: erroresPrevios } = cargarEstadoPrevio(args.out, errPath);
  console.log(`Estado previo: ${resultadosPrevios.length} registros, ${erroresPrevios.length} errores acumulados.`);

  const { pendientes, saltados, fallosPermanentes } = calcularPendientes(groups, resultadosPrevios, erroresPrevios);
  console.log(`${saltados} expedientes ya resueltos (se saltan), ${fallosPermanentes} con fallo permanente (>= ${MAX_INTENTOS} intentos, se saltan), ${pendientes.size} pendientes de procesar. Concurrencia: ${args.concurrency}`);

  // Se descartan los resultados y errores previos de los expedientes que se
  // van a reprocesar, para no duplicar registros ni arrastrar errores viejos
  // ya resueltos (un contrato que fallaba en la corrida anterior y ahora
  // tiene éxito no debe seguir apareciendo en data.errores.json).
  const codigosPendientes = new Set();
  for (const contratos of pendientes.values()) for (const c of contratos) codigosPendientes.add(c.codigoContrato);
  const resultados = resultadosPrevios.filter(r => !codigosPendientes.has(r.codigo_contrato));
  const errores = erroresPrevios.filter(e => !codigosPendientes.has(e.codigoContrato));

  const stats = { preciosCorregidos: [], cantidadesDerivadas: [] };
  const browser = await chromium.launch({ headless: true });

  const queue = [...pendientes.entries()];
  const ctx = { done: 0, total: queue.length };

  const workers = Array.from({ length: args.concurrency }, () =>
    worker(queue, browser, compendio, cucopMap, stats, ctx, resultados, errores, args.out, errPath)
  );
  await Promise.all(workers);

  await browser.close();

  guardarCheckpoint(args.out, errPath, resultados, errores);

  const reporte = construirReporteCalidad(resultados, stats);
  const reportePath = args.out.replace(/\.json$/, '.calidad.json');
  fs.writeFileSync(reportePath, JSON.stringify(reporte, null, 2), 'utf8');

  const excelPath = args.out.replace(/\.json$/, '.xlsx');
  guardarExcel(excelPath, resultados);

  console.log('--- Resumen ---');
  console.log(`Expedientes procesados en esta corrida: ${ctx.total}`);
  console.log(`Registros finales totales: ${resultados.length}`);
  console.log(`Errores totales acumulados: ${errores.length}`);
  console.log(`Precios corregidos por inconsistencia esta corrida: ${stats.preciosCorregidos.length}`);
  console.log(`Cantidades derivadas de subtotal/precio_unitario esta corrida: ${stats.cantidadesDerivadas.length}`);
  console.log(`Salida: ${args.out}`);
  console.log(`Excel: ${excelPath}`);
  console.log(`Errores detallados: ${errPath}`);
  console.log(`Reporte de calidad: ${reportePath}`);
}

main().catch(err => {
  console.error('ERROR FATAL:', err);
  process.exit(1);
});
