// Piezas compartidas entre scripts/extract.js y scripts/validar-claves.js --
// antes vivían duplicadas byte a byte en ambos archivos (con un comentario en
// cada uno pidiendo mantenerlas en sync a mano). Un solo lugar para editar
// evita que las dos copias diverjan.

const XLSX = require('xlsx');

// La clave CSG viene a veces con puntos ("010.000.6153.00") y a veces sin
// ellos ("010000430400") al inicio de una descripción -- se capturan los 4
// grupos de dígitos por separado y se reconstruye siempre en formato con
// puntos. Usada tanto para extraer `clave` de la descripción del contrato
// (extract.js) como, antes, para parsear el catálogo CUCoP+ (ahora
// precomputado una sola vez por scripts/build-cucop.js -- ver ahí).
const CLAVE_RE = /^\s*(\d{3})\.?(\d{3})\.?(\d{4})\.?(\d{2})\s*\.?\s*(.*)$/s;

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

// Descripción de cada columna de las hojas "Cerrado"/"Abierto" del Excel
// (Metodologia.md §4) -- quien reciba data.xlsx suelto, sin este repo a la
// mano, necesita saber qué significa cada campo.
const DICCIONARIO = [
  ['Producto', 'Descripción del medicamento, según el detalle del contrato. Cuando esa descripción viene degenerada (vacía de contenido o sin espacios) se reemplaza por la ficha del catálogo CUCoP+ -- ver Metodologia.md §5.3.2.'],
  ['Compañía', 'Nombre del proveedor o contratista.'],
  ['Fecha de firma', 'Cuándo se formalizó el contrato.'],
  ['Fecha de fallo', 'Cuándo se determinó el precio ganador (adjudicación); normalmente antes de la firma. Vacía en algunas compras consolidadas (ver Limitaciones.md #14).'],
  ['Fecha de inicio', 'Inicio de la ventana de vigencia del contrato.'],
  ['Fecha de fin', 'Fin de la ventana de vigencia del contrato.'],
  ['Precio unitario', 'Precio unitario sin impuestos, validado/recalculado contra el subtotal cuando difieren más de 1% (Metodologia.md §5.3).'],
  ['Precio total (hoja Cerrado)', 'Monto total del contrato (precio unitario × volumen).'],
  ['Precio total mínimo / máximo (hoja Abierto)', 'Piso/techo de exposición contractual (precio unitario × volumen mínimo/máximo).'],
  ['Volumen (hoja Cerrado)', 'Cantidad comprometida en el contrato (no lo entregado realmente).'],
  ['Volumen mínimo / máximo (hoja Abierto)', 'Piso/techo de cantidad comprometida cuando el contrato tiene un rango genuino.'],
  ['Grupo terapéutico', 'Grupo(s) terapéutico(s) asignado(s) vía el Compendio CSG (puede tener más de uno). Vacío si la clave no está en el Compendio.'],
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

// Copia del dataset final en Excel, para quien prefiera trabajarlo ahí en vez
// del JSON que consume el dashboard. Reducido a las columnas de negocio (no
// las de trazabilidad/auditoría del pipeline -- esas quedan en docs/data.json
// para quien las necesite) y partido en dos hojas porque un contrato con
// rango genuino (Abierto) necesita mín/máx por separado, mientras que uno sin
// rango (Cerrado) solo tiene un número real -- mezclarlos en una sola tabla
// dejaría columnas vacías o forzaría a repetir el mismo valor en mín y máx.
function guardarExcel(outPath, resultados) {
  const filasCerrado = [];
  const filasAbierto = [];
  for (const r of resultados) {
    const base = {
      Producto: r.producto,
      'Compañía': r.proveedor,
      'Fecha de firma': r.fecha_firma_contrato,
      'Fecha de fallo': r.fecha_fallo,
      'Fecha de inicio': r.fecha_inicio_contrato,
      'Fecha de fin': r.fecha_fin_contrato,
      'Precio unitario': r.precio_unitario,
    };
    if (esRangoGenuino(r)) {
      filasAbierto.push({
        ...base,
        'Precio total mínimo': r.valor_minimo,
        'Precio total máximo': r.valor_maximo,
        'Volumen mínimo': r.cantidad_minima,
        'Volumen máximo': r.cantidad_maxima,
        'Grupo terapéutico': grupoTerapeuticoTexto(r),
      });
    } else {
      filasCerrado.push({
        ...base,
        'Precio total': r.valor_maximo,
        Volumen: r.cantidad_maxima,
        'Grupo terapéutico': grupoTerapeuticoTexto(r),
      });
    }
  }

  const wsCerrado = XLSX.utils.json_to_sheet(filasCerrado);
  const wsAbierto = XLSX.utils.json_to_sheet(filasAbierto);
  const wsDiccionario = XLSX.utils.aoa_to_sheet([['Campo', 'Descripción'], ...DICCIONARIO]);
  wsDiccionario['!cols'] = [{ wch: 32 }, { wch: 100 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsDiccionario, 'Diccionario');
  XLSX.utils.book_append_sheet(wb, wsCerrado, 'Cerrado');
  XLSX.utils.book_append_sheet(wb, wsAbierto, 'Abierto');
  XLSX.writeFile(wb, outPath);
}

module.exports = { CLAVE_RE, DICCIONARIO, guardarExcel, dividirProductoSiAplica };
