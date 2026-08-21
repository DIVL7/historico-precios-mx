// Piezas compartidas entre scripts/extract.js y scripts/validar-claves.js --
// antes vivían duplicadas byte a byte en ambos archivos (con un comentario en
// cada uno pidiendo mantenerlas en sync a mano). Un solo lugar para editar
// evita que las dos copias diverjan.

const ExcelJS = require('exceljs');

// La clave CSG viene a veces con puntos ("010.000.6153.00") y a veces sin
// ellos ("010000430400") al inicio de una descripción -- se capturan los 4
// grupos de dígitos por separado y se reconstruye siempre en formato con
// puntos. Usada tanto para extraer `clave` de la descripción del contrato
// (extract.js) como, antes, para parsear el catálogo CUCoP+ (ahora
// precomputado una sola vez por scripts/build-cucop.js -- ver ahí).
const CLAVE_RE = /^\s*(\d{3})\.?(\d{3})\.?(\d{4})\.?(\d{2})\s*\.?\s*(.*)$/s;

// Devuelve el texto de `s` después del prefijo de clave CSG si CLAVE_RE
// matcheó, o `s` tal cual si no -- centralizado aquí (no reimplementado en
// extract.js/validar-claves.js) por la misma razón que el resto de este
// archivo: validar-claves.js lo usa para no confundir los dígitos de la
// clave con dosis/cantidad al hacer matching numérico contra el catálogo
// (ver extraerNumerosProducto ahí), sin duplicar el match() de CLAVE_RE.
function quitarPrefijoClave(s) {
  const m = String(s || '').match(CLAVE_RE);
  return m ? m[5] : s;
}

// Casos puntuales confirmados a mano donde `producto` en realidad lista
// varios ingredientes activos separados por " - " y cada uno debería quedar
// como su propio registro (mismos datos -- clave, precio, cantidad, fechas
// -- solo cambia el texto de `producto`). NO es una regla general: la
// mayoría de los " - " en este dataset son parte de rangos numéricos
// ("LONGITUD: 17 - 24 MM"), tablas de ingredientes, o nombres de
// combinación con la dosis pegada solo al último segmento ("TELMISARTAN -
// HIDROCLOROTIAZIDA TABLETA 80.0 MG/12.5 MG 14 TABLETAS" -- splitear ahí
// deja "TELMISARTAN" sin dosis ni presentación). Se probó una heurística
// automática (segmentos cortos, sin números) y de 445 productos únicos con
// "-" en el dataset real, apenas 16 calificaban -- y ni esos 16 eran
// confiables (ej. "CLORANFENICOL - TUBO APLICADOR" no son dos medicamentos).
// Verificado el 2026-08-17: se agregan acá uno por uno, a mano, solo cuando
// se confirma contra el texto real del contrato que de verdad son
// ingredientes independientes sin nada más pegado.
const PRODUCTOS_A_DIVIDIR = new Set([
  'BECLOMETASONA - FORMOTEROL - GLICOPIRRONIO',
]);

// Devuelve un array de nombres de producto: [producto] sin cambios si no es
// uno de los casos confirmados en PRODUCTOS_A_DIVIDIR, o los segmentos
// divididos si sí lo es.
function dividirProductoSiAplica(producto) {
  if (!PRODUCTOS_A_DIVIDIR.has(producto)) return [producto];
  return producto.split(/\s*-\s*/).map(p => p.trim()).filter(Boolean);
}

// Descripción de cada columna de las hojas "Cerrado"/"Abierto" del Excel --
// quien reciba data.xlsx suelto, sin este repo a la mano, necesita saber qué
// significa cada campo sin tener que ir a buscar otro archivo. Donde el
// campo corresponde a una columna del "Diccionario de datos del Reporte de
// Datos Relevantes del Contrato" publicado por CompraNet
// (DD_PIC_CONTRATOS_2400703), se usa esa definición oficial; el resto son
// campos calculados o ajenos al reporte de CompraNet y llevan una
// definición propia.
const DICCIONARIO = [
  ['PRODUCTO', 'Nombre y presentación del medicamento comprado, según el detalle del contrato.'],
  ['COMPAÑÍA', 'Razón social o nombre de la persona física o moral que celebra el contrato con la dependencia o entidad.'],
  ['INSTITUCIÓN', 'Siglas que identifican a la dependencia o entidad que compra.'],
  ['FECHA DE INICIO', 'Fecha en que inicia la vigencia del contrato.'],
  ['FECHA DE FIN', 'Fecha en que termina la vigencia del contrato.'],
  ['MONEDA', 'Código de la moneda en que se pactó el contrato (por ejemplo MXN, USD). La inmensa mayoría del dataset es MXN, con un puñado de contratos en otras monedas -- el precio unitario y el precio total de esa misma fila llevan el símbolo de esta moneda, no siempre "$" de pesos.'],
  ['PRECIO UNITARIO', 'Precio de una sola unidad del medicamento, sin impuestos.'],
  ['PRECIO TOTAL (hoja Cerrado)', 'Monto total del contrato: precio unitario multiplicado por el volumen.'],
  ['PRECIO TOTAL MÍNIMO / MÁXIMO (hoja Abierto)', 'Monto mínimo y máximo que puede llegar a pagarse por el contrato: precio unitario multiplicado por el volumen mínimo y máximo.'],
  ['VOLUMEN (hoja Cerrado)', 'Cantidad de unidades del medicamento que compromete el contrato (no necesariamente lo entregado realmente).'],
  ['VOLUMEN MÍNIMO / MÁXIMO (hoja Abierto)', 'Cantidad mínima y máxima de unidades que compromete el contrato, cuando el volumen no es un número fijo sino un rango.'],
  ['GRUPO TERAPÉUTICO', 'Categoría médica del medicamento (por ejemplo, analgésicos, antibióticos); puede tener más de una. Queda vacío si el medicamento no aparece en el catálogo oficial de medicamentos usado para clasificarlo.'],
];

// Un contrato va a la hoja "Abierto" cuando su volumen Y su precio total
// (las dos, no basta con una sola) tienen un rango real (mín != máx) -- NO
// se usa la etiqueta `tipo_contrato` (viene tal cual del CSV oficial,
// columna "Tipo de contrato", y ~29% del dataset la trae vacía porque la
// fuente la dejó en blanco; además no siempre coincide con si el rango
// numérico es genuino). Se decidió así el 2026-08-17: la regla se basa en
// los números reales de cada registro, no en una etiqueta ajena que puede
// faltar o no reflejar el dato. AND, no OR (corregido el mismo día): si solo
// una de las dos varía, no alcanza para considerarlo un rango genuino.
//
// Excepción: precio_unitario === 0 hace que valor_minimo/valor_maximo den 0
// en los dos extremos sin importar la cantidad real (0 × lo que sea = 0) --
// ahí el valor no aporta ninguna señal, así que se evalúa solo la cantidad.
// Verificado el 2026-08-17 contra el dataset real: sin esta excepción, 3
// registros con precio_unitario=0 y rango de cantidad genuino perdían
// silenciosamente el mínimo en la hoja "Cerrado" (solo se veía el máximo).
function esRangoGenuino(r) {
  if (r.precio_unitario === 0) return r.cantidad_minima !== r.cantidad_maxima;
  return r.cantidad_minima !== r.cantidad_maxima && r.valor_minimo !== r.valor_maximo;
}

function grupoTerapeuticoTexto(r) {
  return Array.isArray(r.grupo_terapeutico) ? r.grupo_terapeutico.join(', ') : r.grupo_terapeutico;
}

// Símbolo real a anteponer en el formato de celda de las columnas de precio,
// por código de moneda -- el dataset es ~99.99% MXN pero trae un puñado de
// contratos en otras monedas (ver columna MONEDA). Se usa el símbolo propio
// de cada moneda tal cual (MXN y USD comparten "$"); la columna MONEDA de al
// lado es la que desambigua esos casos, no el símbolo. Si aparece un código
// no listado aquí, se antepone el código tal cual (ej. "MWK 1,234.00") en vez
// de asumir "$" a ciegas.
const SIMBOLOS_MONEDA = { MXN: '$', USD: '$', EUR: '€', CAD: '$', GBP: '£', JPY: '¥' };

// El código de moneda no listado arriba se interpola tal cual dentro de un
// literal entre comillas dobles del numFmt de Excel -- una comilla doble (o
// una barra invertida, que ExcelJS interpreta como escape) en el valor crudo
// del CSV cerraría ese literal antes de tiempo y rompería el formato de la
// celda. Un código ISO de moneda real nunca las trae; se quitan por si acaso
// llega un valor mal capturado en la fuente.
function sanitizarParaNumFmt(s) {
  return String(s).replace(/["\\]/g, '');
}

function formatoMoneda(moneda) {
  const simbolo = SIMBOLOS_MONEDA[moneda] || (moneda ? `${sanitizarParaNumFmt(moneda)} ` : '$');
  return `"${simbolo}"#,##0.00`;
}

// Solo comas de miles, sin decimales ni símbolo -- para columnas de volumen
// (unidades de medicamento, no dinero). Sin decimales porque `cantidad_minima`
// / `cantidad_maxima` son siempre enteras en la práctica (verificado: 0 de
// 47,725 registros trae un valor no entero en ninguna de las dos), así que
// dos decimales solo agregarían ".00" sin aportar nada.
const FORMATO_VOLUMEN = '#,##0';

// Aplica el formato de celda (número visual, NO trunca el valor guardado --
// decisión explícita del usuario 2026-08-21: quien abra el Excel y necesite
// más precisión la sigue teniendo en el valor real, solo la vista redondea a
// 2 decimales) a las columnas de precio (con símbolo de MONEDA por fila) y de
// volumen (sin símbolo). Recibe las `Row` que ya devolvió `ws.addRows(filas)`
// -- en el mismo orden que `filas` -- en vez de volver a buscarlas por índice
// con `ws.getRow()`, que ya las tiene resueltas.
function aplicarFormatoNumerico(rows, filas, columnasPrecio, columnasVolumen) {
  rows.forEach((row, i) => {
    const fmtPrecio = formatoMoneda(filas[i].MONEDA);
    for (const col of columnasPrecio) row.getCell(col).numFmt = fmtPrecio;
    for (const col of columnasVolumen) row.getCell(col).numFmt = FORMATO_VOLUMEN;
  });
}

// Borde fino gris claro, el mismo en las tres hojas -- da la cuadrícula de
// "tabla" sin ser tan marcado como un borde negro por defecto de Excel.
const BORDE_FINO = { style: 'thin', color: { argb: 'FFB7B7B7' } };
const BORDES_CELDA = { top: BORDE_FINO, left: BORDE_FINO, bottom: BORDE_FINO, right: BORDE_FINO };

// Da a una hoja el mismo formato simple en las tres: encabezado en negrita
// sobre fondo gris, fila 1 fija al hacer scroll, y borde fino en cada celda.
// `columnas` es siempre un array de columnas ExcelJS completas
// (`{header, key, width}`) -- ver columnasDesdeClaves() para el caso
// Cerrado/Abierto (ancho ajustado al título) y guardarExcel() para
// Diccionario (ancho fijo propio, key distinto del header).
function formatearHoja(ws, columnas) {
  ws.columns = columnas.map(col => ({ style: { border: BORDES_CELDA }, ...col }));
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.getRow(1).eachCell(cell => {
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
  });
}

// Columnas ExcelJS derivadas de un array de nombres tal cual (Cerrado/
// Abierto): header y key son el mismo texto, ancho ajustado al título (no al
// contenido -- un producto con nombre kilométrico no debe estirar la
// columna).
function columnasDesdeClaves(keys) {
  return keys.map(k => ({ header: k, key: k, width: k.length + 4 }));
}

// Copia del dataset final en Excel, para quien prefiera trabajarlo ahí en vez
// del JSON que consume el dashboard. Reducido a las columnas de negocio (no
// las de trazabilidad/auditoría del pipeline -- esas quedan en docs/data.json
// para quien las necesite) y partido en dos hojas porque un contrato con
// rango genuino (Abierto) necesita mín/máx por separado, mientras que uno sin
// rango (Cerrado) solo tiene un número real -- mezclarlos en una sola tabla
// dejaría columnas vacías o forzaría a repetir el mismo valor en mín y máx.
async function guardarExcel(outPath, resultados) {
  const filasCerrado = [];
  const filasAbierto = [];
  for (const r of resultados) {
    const base = {
      PRODUCTO: r.producto,
      'COMPAÑÍA': r.proveedor,
      'INSTITUCIÓN': r.institucion,
      'FECHA DE INICIO': r.fecha_inicio_contrato,
      'FECHA DE FIN': r.fecha_fin_contrato,
      MONEDA: r.moneda || '',
      'PRECIO UNITARIO': r.precio_unitario,
    };
    if (esRangoGenuino(r)) {
      filasAbierto.push({
        ...base,
        'PRECIO TOTAL MÍNIMO': r.valor_minimo,
        'PRECIO TOTAL MÁXIMO': r.valor_maximo,
        'VOLUMEN MÍNIMO': r.cantidad_minima,
        'VOLUMEN MÁXIMO': r.cantidad_maxima,
        'GRUPO TERAPÉUTICO': grupoTerapeuticoTexto(r),
      });
    } else {
      filasCerrado.push({
        ...base,
        'PRECIO TOTAL': r.valor_maximo,
        VOLUMEN: r.cantidad_maxima,
        'GRUPO TERAPÉUTICO': grupoTerapeuticoTexto(r),
      });
    }
  }

  const wb = new ExcelJS.Workbook();

  const wsDiccionario = wb.addWorksheet('Diccionario');
  // Esta hoja es de consulta (definiciones largas), no datos tabulares --
  // ancho fijo generoso en vez de ajustado al título, si no "DESCRIPCIÓN"
  // quedaría ilegible.
  formatearHoja(wsDiccionario, [
    { header: 'CAMPO', key: 'campo', width: 32 },
    { header: 'DESCRIPCIÓN', key: 'descripcion', width: 100 },
  ]);
  wsDiccionario.addRows(DICCIONARIO.map(([campo, descripcion]) => ({ campo, descripcion })));

  const wsCerrado = wb.addWorksheet('Cerrado');
  formatearHoja(wsCerrado, columnasDesdeClaves(Object.keys(filasCerrado[0] || {})));
  const rowsCerrado = wsCerrado.addRows(filasCerrado);
  aplicarFormatoNumerico(rowsCerrado, filasCerrado, ['PRECIO UNITARIO', 'PRECIO TOTAL'], ['VOLUMEN']);

  const wsAbierto = wb.addWorksheet('Abierto');
  formatearHoja(wsAbierto, columnasDesdeClaves(Object.keys(filasAbierto[0] || {})));
  const rowsAbierto = wsAbierto.addRows(filasAbierto);
  aplicarFormatoNumerico(
    rowsAbierto,
    filasAbierto,
    ['PRECIO UNITARIO', 'PRECIO TOTAL MÍNIMO', 'PRECIO TOTAL MÁXIMO'],
    ['VOLUMEN MÍNIMO', 'VOLUMEN MÁXIMO'],
  );

  await wb.xlsx.writeFile(outPath);
}

module.exports = { CLAVE_RE, DICCIONARIO, guardarExcel, dividirProductoSiAplica, quitarPrefijoClave };
