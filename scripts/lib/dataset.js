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

// Descripción de cada columna de la hoja "Precios" del Excel (Metodologia.md
// §4) -- quien reciba data.xlsx suelto, sin este repo a la mano, necesita
// saber qué significa cada campo. Los nombres coinciden con docs/data.json
// salvo cantidad_min/cantidad_max, que en el Excel reemplazan a
// cantidad_minima/cantidad_maxima (mismo campo, nombre abreviado -- ver
// guardarExcel).
const DICCIONARIO = [
  ['codigo_contrato', 'Identificador único del contrato en Compras MX (ej. C-2026-000123).'],
  ['num_contrato', 'Número de contrato asignado por la institución compradora (formato varía por institución).'],
  ['clave', 'Clave del Compendio Nacional de Medicamentos (CSG), formato NNN.NNN.NNNN.NN. Vacía si no se pudo determinar.'],
  ['clave_fuente', 'De dónde salió "clave": descripcion, cucop, validacion_nombre_local, propagacion_confiable, o vacío si no se pudo determinar.'],
  ['cve_cucop', 'Clave del catálogo CUCoP+ reportada por la institución compradora.'],
  ['producto', 'Descripción del medicamento, según el detalle del contrato. Cuando esa descripción viene degenerada (vacía de contenido o sin espacios) se reemplaza por la ficha del catálogo CUCoP+ -- ver Metodologia.md §5.3.2. Prefijos de numeración/viñetas al inicio se recortan -- ver §5.3.3. Coletilla administrativa al final ("CONFORME A PARTIDA...") se recorta -- ver §5.3.4.'],
  ['tipo_insumo', 'Siempre MEDICAMENTO -- es el alcance de esta base.'],
  ['grupo_terapeutico', 'Grupo(s) terapéutico(s) asignado(s) vía el Compendio CSG (puede tener más de uno). Vacío si la clave no está en el Compendio.'],
  ['procedimiento', 'Número de procedimiento de contratación.'],
  ['tipo_procedimiento', 'CONSOLIDADA (varias instituciones agrupadas en un solo procedimiento) o NO CONSOLIDADA (una sola institución).'],
  ['tipo_contrato', 'Abierto (rango mín-máx) o Cerrado (cantidad fija), valor oficial de la fuente.'],
  ['proveedor', 'Nombre del proveedor o contratista.'],
  ['institucion', 'Siglas de la institución compradora. En compras consolidadas puede no ser la institución que realmente contrató (ver Limitaciones.md).'],
  ['unidad_medida', 'Unidad de medida de la cantidad (pieza, kilogramo, etc.).'],
  ['precio_unitario', 'Precio unitario sin impuestos, validado/recalculado contra el subtotal cuando difieren más de 1% (Metodologia.md §5.3).'],
  ['cantidad_min', 'Cantidad mínima comprometida en el contrato (no lo entregado realmente). Igual a cantidad_max cuando el contrato no tiene rango genuino -- Cerrado, o "por monto" derivado de subtotal/precio_unitario -- y solo difiere de cantidad_max en contratos Abierto con rango real.'],
  ['cantidad_max', 'Cantidad máxima comprometida en el contrato. Igual a cantidad_min cuando no hay rango genuino (ver cantidad_min).'],
  ['valor_minimo', 'Piso de exposición contractual garantizado (precio_unitario × cantidad_min, o subtotal directo). Igual a valor_maximo cuando no hay rango genuino.'],
  ['valor_maximo', 'Techo de exposición contractual posible (precio_unitario × cantidad_max, o subtotal directo). Igual a valor_minimo cuando no hay rango genuino -- usar cualquiera de los dos (nunca sumar ambos) para sumas/promedios agregados (Metodologia.md §5.5).'],
  ['fecha_firma_contrato', 'Cuándo se formalizó el contrato.'],
  ['fecha_fallo', 'Cuándo se determinó el precio ganador (adjudicación); normalmente antes de la firma. Vacía en algunas compras consolidadas (ver Limitaciones.md #14).'],
  ['fecha_inicio_contrato', 'Inicio de la ventana de vigencia del contrato.'],
  ['fecha_fin_contrato', 'Fin de la ventana de vigencia del contrato.'],
];

// Copia del dataset final en Excel, para quien prefiera trabajarlo ahí en vez
// del JSON que consume el dashboard. grupo_terapeutico es un arreglo en el
// JSON (un medicamento puede pertenecer a más de un grupo) -- en Excel se
// aplana a texto separado por coma, porque una celda no puede guardar un
// arreglo.
//
// cantidad_minima/cantidad_maxima y valor_minimo/valor_maximo ya vienen
// siempre poblados en docs/data.json (iguales entre sí cuando el contrato no
// tiene rango genuino, ver extract.js buildRegistro) -- no hace falta
// colapsarlos a una sola columna aquí. Solo se renombran cantidad_minima/
// cantidad_maxima a cantidad_min/cantidad_max (más cortos) para las columnas
// del Excel; no hay columna `valor` suelta porque sería idéntica a
// valor_maximo en el 100% de los casos (verificado contra el dataset
// completo -- `cantidad` siempre usaba cantidad_maxima cuando había una
// disponible).
function guardarExcel(outPath, resultados) {
  const filas = resultados.map(r => ({
    codigo_contrato: r.codigo_contrato,
    num_contrato: r.num_contrato,
    clave: r.clave,
    clave_fuente: r.clave_fuente,
    cve_cucop: r.cve_cucop,
    producto: r.producto,
    tipo_insumo: r.tipo_insumo,
    grupo_terapeutico: Array.isArray(r.grupo_terapeutico) ? r.grupo_terapeutico.join(', ') : r.grupo_terapeutico,
    procedimiento: r.procedimiento,
    tipo_procedimiento: r.tipo_procedimiento,
    tipo_contrato: r.tipo_contrato,
    proveedor: r.proveedor,
    institucion: r.institucion,
    unidad_medida: r.unidad_medida,
    precio_unitario: r.precio_unitario,
    cantidad_min: r.cantidad_minima,
    cantidad_max: r.cantidad_maxima,
    valor_minimo: r.valor_minimo,
    valor_maximo: r.valor_maximo,
    fecha_firma_contrato: r.fecha_firma_contrato,
    fecha_fallo: r.fecha_fallo,
    fecha_inicio_contrato: r.fecha_inicio_contrato,
    fecha_fin_contrato: r.fecha_fin_contrato,
  }));
  const ws = XLSX.utils.json_to_sheet(filas);
  const wsDiccionario = XLSX.utils.aoa_to_sheet([['Campo', 'Descripción'], ...DICCIONARIO]);
  wsDiccionario['!cols'] = [{ wch: 22 }, { wch: 100 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsDiccionario, 'Diccionario');
  XLSX.utils.book_append_sheet(wb, ws, 'Precios');
  XLSX.writeFile(wb, outPath);
}

module.exports = { CLAVE_RE, DICCIONARIO, guardarExcel };
