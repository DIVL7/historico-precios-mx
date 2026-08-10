# Metodología: Base de Datos Histórica de Precios de Medicamentos — Sector Salud México

**Fuente primaria:** Compras MX (Secretaría Anticorrupción y Buen Gobierno)
**Fuente de referencia:** Compendio Nacional de Insumos para la Salud — Medicamentos (Consejo de Salubridad General, "CNIS")

**Antes de usar los datos, leer `Limitaciones.md`** — documenta dónde puede haber datos erróneos o incompletos y las alternativas disponibles para cada caso. Este documento describe cómo funciona la herramienta; no repite ese contenido.

---

## 1. Objetivo

Construir una base de datos histórica y consultable de precios unitarios de medicamentos adquiridos por el sector salud público mexicano, a partir de los registros de contrataciones públicas publicados en Compras MX, cruzados contra el catálogo oficial de medicamentos del CSG.

---

## 2. Alcance

### 2.1 Producto

**Únicamente medicamentos.** En el clasificador CUCoP (Clasificador Único de Contrataciones Públicas), los medicamentos caen exclusivamente bajo la **partida específica 25301** (partida genérica 2530, "Medicinas y productos farmacéuticos") — cubre el 100% del universo de medicamentos, incluyendo vacunas y biológicos.

Otras categorías de "producto de salud" (material de curación, instrumental, equipo médico y de laboratorio) están fuera de alcance: Compras MX no incrusta una clave del compendio oficial en la descripción de esos artículos, lo que impediría validar de forma confiable si un artículo pertenece a un catálogo oficial o es un genérico sin ficha.

### 2.2 Horizonte temporal

Los archivos de "Contratos de la Plataforma Integral" en Compras MX cubren de forma consistente **2020 a 2026** con un esquema homogéneo. Años anteriores (CompraNet 5.0: ~2010–2022; CompraNet 3.0: 2002–2011) usan esquemas de datos distintos y más limitados (sin desglose de partidas ni claves CUCoP+ por línea) — fuera de alcance de esta herramienta.

### 2.3 Fuentes utilizadas

| Fuente | Rol |
|---|---|
| Compras MX — Datos Abiertos (CSV masivo anual) | Descubrimiento y filtrado de contratos relevantes |
| Compras MX — API interna (`detallepartidas`) | Extracción de precio unitario, clave y cantidad por partida |
| Compendio Nacional de Medicamentos (CSG) | Catálogo de validación + fuente de `grupo_terapéutico` |
| Catálogo CUCoP+ (`cucop.xlsx`) | Respaldo para recuperar la clave CSG cuando la institución no la cita |

---

## 3. Arquitectura técnica de las fuentes

### 3.1 Datos Abiertos de Compras MX

Portal: `https://comprasmx.buengobierno.gob.mx/datos-abiertos`. Archivo: **"Contratos de la Plataforma Integral Compras MX / CompraNet [año]"**, uno por año. Cada fila es un contrato individual (no un expediente completo). Trae todas las columnas usadas en el modelo de datos (§4) más `Dirección del anuncio` — URL al detalle del expediente, incluye el hash necesario para la extracción.

El archivo se sirve vía un formulario protegido con reCAPTCHA v3 — la descarga requiere una sesión de navegador real, no es una URL pública.

### 3.2 API interna y reCAPTCHA v3

Endpoints (dominio `upcp-cnetservicios.buengobierno.gob.mx`, backend "whitney"):

```
GET  /whitney/sitiopublico/expedientes/{hash}?id_proceso=procedimiento
GET  /whitney/sitiopublico/detallepartidas/{hash}/{codigo_contrato}
```

Ambos requieren headers de reCAPTCHA v3 (`grc`, `igrc`, `xgrc`) generados dinámicamente en el navegador — no responden a peticiones HTTP directas (`curl`). La extracción usa automatización de navegador real (Playwright) para generar estos headers de forma legítima.

`detallepartidas` entrega el precio unitario y cantidad por partida, en uno de dos esquemas de contrato:
- **Abierto** (rango): `Cantidad mínima`, `Cantidad máxima`, precio unitario, montos por cantidad mínima/máxima.
- **Cantidad fija**: `Cantidad solicitada`, precio unitario, subtotal, IVA, total.

### 3.3 Compendio Nacional de Medicamentos (CSG)

Fuente: `https://csg.gob.mx/COMPENDIO_MEDICAMENTOS_V07042026.xlsm`. Columnas: `Grupo` (grupo terapéutico, puede haber más de uno por clave), `Clave` (formato `NNN.NNN.NNNN.NN`), `Insumo`, `Descripción`, `Indicaciones`. Se usa como tabla de referencia para asignar `grupo_terapéutico` a cada renglón.

La celda `Grupo` trae cada grupo prefijado con un marcador tipo `Grupo Nº 23:` (número de grupo, estable) seguido del nombre en texto libre — pero el separador entre grupos y la ortografía del nombre libre no son consistentes en el archivo fuente (a veces sin acentos, a veces con erratas de mayúsculas, a veces sin salto de línea entre dos grupos). `scripts/build-compendio.js` parte la celda por el marcador (no por saltos de línea) y traduce el número a un nombre canónico fijo (`GRUPOS_TERAPEUTICOS`, 23 grupos) en vez de usar el texto libre tal cual — así una errata en el nombre no genera un grupo terapéutico duplicado. Si el CSG agrega un grupo nuevo (número fuera del catálogo conocido), el build imprime una advertencia y usa el texto libre de respaldo en vez de fallar o inventar un nombre.

### 3.4 Catálogo CUCoP+

Fuente: `data/raw/cucop.xlsx`. Para toda entrada de la partida 25301, la columna `DESCRIPCIÓN` trae la clave del Compendio como prefijo — es estático (no depende de qué institución compre), así que sirve de respaldo cuando la descripción libre del contrato no incluye la clave directamente.

Igual que el Compendio (§3.3), el xlsx crudo no se versiona (`.gitignore`) y se convierte una sola vez a un JSON liviano indexado por `cve_cucop` — `scripts/build-cucop.js` genera `data/cucop_medicamentos.json`, que sí se commitea. `extract.js` y `validar-claves.js` solo leen ese JSON en runtime; ninguno de los dos necesita el xlsx crudo, lo cual importa en particular para CI (el workflow de GitHub Actions nunca descarga ni recibe ese archivo).

---

## 4. Modelo de datos

| Campo | Fuente | Notas |
|---|---|---|
| `codigo_contrato` | Compras MX | Identificador único a nivel contrato |
| `num_contrato` | Compras MX ("Núm. del contrato") | Número de contrato asignado por la institución compradora — formato libre, distinto de `codigo_contrato` |
| `clave` | Ver §5.2 (asignación y validación de clave) | Formato CSG `NNN.NNN.NNNN.NN`, o `null` si no se pudo determinar |
| `clave_fuente` | Interno | De dónde salió `clave`: `descripcion` (del texto del contrato), `cucop` (vía catálogo CUCoP), `validacion_nombre_local` (por nombre contra el Compendio), `propagacion_confiable` (heredada de otro registro del mismo producto), o `null` |
| `producto` | "Descripción detallada" (`detallepartidas`), texto restante tras remover la clave, salvo cuando viene degenerado — ahí se usa la ficha CUCoP+ en su lugar, ver §5.3.2. Prefijos de numeración/viñetas al inicio se recortan, ver §5.3.3 | |
| `tipo_insumo` | Constante `MEDICAMENTO` | |
| `grupo_terapéutico` | Join contra Compendio CSG por `clave` | `null` si `clave` es `null` o no existe en el Compendio |
| `procedimiento` | "Número de procedimiento" | |
| `tipo_procedimiento` | "Compra consolidada" SI/NO → CONSOLIDADA / NO CONSOLIDADA | Consolidada: varias instituciones agrupan sus necesidades en un solo procedimiento (coordinado por una entidad líder) en vez de licitar cada una por separado. No consolidada: una institución licita y contrata por su cuenta |
| `proveedor` | "Proveedor o contratista" | |
| `institución` | "Siglas de la Institución" (CSV) | En compras consolidadas puede no corresponder a la institución que realmente contrató cada renglón — ver `Limitaciones.md` |
| `precio_unitario` | "Precio unitario sin impuestos" (`detallepartidas`), recalculado si es inconsistente con el subtotal — ver §5.3 | |
| `cantidad_minima` / `cantidad_maxima` | "Cantidad mínima" / "Cantidad máxima" (`detallepartidas`) cuando hay rango genuino; si no, ambas colapsan al mismo valor único (`cantidad_maxima ?? cantidad ?? cantidad_minima` de la fuente, o derivado de `subtotal / precio_unitario` — ver §5.3.1) | Siempre pobladas con el mismo valor entre sí, salvo en contratos "Abierto" con rango real (divergen) o servicios/medicina magistral sin ninguna cantidad reportada (`null` en ambas — ver `Limitaciones.md`) |
| `tipo_contrato` | "Tipo de contrato" (CSV) | `Abierto` / `Cerrado`, valor oficial de la fuente — no se infiere de si `cantidad_maxima` difiere de `cantidad_minima`, para tener una sola fuente de verdad sin ambigüedad al integrar ambos esquemas de cantidad en la misma tabla |
| `valor_minimo` / `valor_maximo` | `precio_unitario × cantidad_minima` / `precio_unitario × cantidad_maxima`, o `subtotal` directo si no se puede calcular así — ver §5.3.1 | Piso y techo de exposición contractual, siempre poblados y coincidentes entre sí salvo en "Abierto" con rango real — ver §5.5 |
| `fecha_firma_contrato` | "Fecha firma contrato" (CSV) | Cuándo se formalizó el contrato |
| `fecha_fallo` | "Fecha de fallo" (CSV) | Cuándo se determinó el precio ganador (adjudicación) — normalmente precede a la firma |
| `fecha_inicio_contrato` / `fecha_fin_contrato` | "Fecha de inicio del contrato" / "Fecha de fin del contrato" (CSV) | Ventana de vigencia del contrato — cuándo ese precio estuvo activo para entregas, distinto de cuándo se firmó |

---

## 5. Pipeline

### 5.1 Extracción (`scripts/extract.js`)

1. Descargar el CSV anual desde Datos Abiertos (`scripts/download-csv.js`, vía navegador automatizado).
2. Filtrar filas donde `Partida específica` contiene `25301` — este filtro opera a nivel de **contrato** (una fila del CSV masivo), no de ítem individual.
3. Agrupar filas por expediente (hash extraído de `Dirección del anuncio`).
4. Por cada expediente, navegar al detalle y localizar cada contrato dentro de la tabla paginada de resultados.
5. Extraer `detallepartidas/{hash}/{codigo_contrato}` — la respuesta JSON trae los ítems ya estructurados (`cve_cucop`, `descripcion`, `precio_unitario`, `subtotal`, `cantidad`, `cantidad_minima`, `cantidad_maxima`, `um`). Un contrato que pasó el filtro del paso 2 puede traer aquí líneas de **otras partidas** (compra mixta: radiofármacos, material de curación, papelería, etc.) — se descarta cualquier ítem cuyo `cve_cucop` no empiece con `25301`, para no marcarlo como medicamento solo por venir en el mismo contrato. (Detectado el 2026-08-09: sin este segundo filtro, ~32% del dataset — 7,989 de 24,909 registros — eran contaminación de otras partidas; se limpió retroactivamente y se corrigió el pipeline.)
6. Por cada ítem que pasa el filtro anterior, asignar clave (§5.2), sanear `producto` si viene degenerado (§5.3.2), calcular `precio_unitario` corregido (§5.3) y `valor_minimo`/`valor_maximo` (§5.5), y construir el registro final.
7. Escribir a `docs/data.json` con checkpoint incremental (reanudable: si el proceso se corta, la siguiente corrida retoma desde el último punto guardado en vez de reprocesar todo).

### 5.2 Asignación y validación de clave

**Durante la extracción**, cada ítem recibe una clave por el método más rápido disponible, en orden:
1. Regex sobre la descripción del contrato (`\d{3}\.\d{3}\.\d{4}\.\d{2}`) — solo algunas instituciones citan la clave directamente.
2. Si no está, recuperación vía `cve_cucop` contra el catálogo CUCoP+.

**Después de la extracción** (`scripts/validar-claves.js`), corre un sanity check sobre el dataset completo: agrupa todos los registros por producto canónico (nombre + huella numérica de dosis/envase, no texto exacto) y determina una única clave "ganadora" por grupo, en orden de confianza:
1. Cualquier miembro con clave extraída directo de la descripción del contrato (la fuente más confiable).
2. Cualquier miembro con clave vía CUCoP que **autoverifica**: la dosis/presentación que ese código declara en el catálogo CUCoP+ debe coincidir (≥2 números) con lo que el contrato realmente describe.
3. Búsqueda por nombre contra el Compendio Nacional local.
4. Si nada de lo anterior aplica, cualquier clave vía CUCoP que no haya autoverificado se **vacía explícitamente** — no se deja un valor que ya se sabe incorrecto solo por falta de reemplazo.

La clave ganadora de cada grupo se propaga a todas sus instancias, así el mismo producto no puede terminar con dos claves distintas según qué contrato lo haya traído. Todo lo que no se resuelve queda documentado en `docs/data.correcciones.json`, sin inventar una clave sin evidencia. Ver `Limitaciones.md` para la cobertura resultante y las alternativas para lo no resuelto (incluye búsqueda externa contra `vademecum.es/cnis`, evaluada pero no automatizada).

### 5.3 Corrección de `precio_unitario`

El subtotal reportado por la API es consistente con `cantidad_minima` en todos los casos verificados. El pipeline recalcula `precio_unitario = subtotal / cantidad_minima` y solo conserva el valor original de la API cuando coincide (±1%); cada corrección queda registrada en `docs/data.calidad.json`.

### 5.3.1 Derivación de `cantidad_minima`/`cantidad_maxima` y `valor_minimo`/`valor_maximo` en contratos "por monto"

Un subconjunto de ítems trae `tipo_contrato_abierto: "MONTO"` en la respuesta de `detallepartidas`: el compromiso contractual es un techo de gasto, no una cantidad de piezas, así que `cantidad_minima`/`cantidad_maxima` (y la cantidad de origen que las respalda) vienen vacíos desde el origen (no es un hueco de la extracción). Sí trae `subtotal` — el "Monto de la Oferta" visible en el detalle del contrato en el sitio — que permite derivar dos cosas:

- **`cantidad_minima`/`cantidad_maxima`** (informativa, ambas colapsan al mismo valor): `subtotal / precio_unitario`, redondeada al entero más cercano, solo cuando las vías directas fallan. Cada caso queda registrado en `docs/data.calidad.json` (`cantidades_derivadas_de_subtotal_entre_precio_unitario`).
- **`valor_minimo`/`valor_maximo`**: se toman de `subtotal` directo, no de `precio_unitario × cantidad_derivada` — multiplicar de vuelta sería un viaje de ida y vuelta innecesario que puede perder precisión por el redondeo intermedio de la cantidad derivada. Este respaldo directo aplica siempre que no haya un rango `cantidad_minima`/`cantidad_maxima` genuino que preservar (ver §5.5), incluyendo el caso donde `precio_unitario` viene en `0` desde el origen (visto en servicios y medicina magistral): ahí no se puede derivar ninguna cantidad tampoco (`cantidad_minima`/`cantidad_maxima` quedan en `null`, único caso donde eso ocurre — ver `Limitaciones.md`), pero `valor_minimo`/`valor_maximo` sí se rescatan de `subtotal`.

### 5.3.2 Reemplazo de `producto` degenerado con la ficha CUCoP+

La institución compradora a veces captura "Descripción detallada" (el campo que alimenta `producto`) de forma degenerada, verificado en vivo contra el sitio en dos formas distintas:

- **Referencia vacía de contenido**: el texto es literalmente `CONFORME A PARTIDA N DE LA CONVOCATORIA` — una remisión a su propia partida, sin describir el producto (visto en el contrato `C-2026-00073692`).
- **Sin espacios entre palabras**: el texto sí describe el producto pero viene concatenado (`ACICLOVIR200MGENVASECON25COMPRIMIDOSOTABLETAS...`) — de origen, no un artefacto de la extracción; se confirmó que los 174 ítems de un mismo procedimiento (`IA-13-312-013000999-T-327-2026`) vienen así (contrato `C-2026-00070247`).

En ambos casos, la página del expediente muestra una columna separada, "Descripción CUCoP+", que siempre viene bien formada — es la ficha oficial del catálogo, no depende de cómo la institución tecleó su propia descripción. El pipeline detecta ambos patrones (regex para el primero, longitud + ausencia de espacios para el segundo) y sustituye `producto` por `cucopMap[cve_cucop].descripcion` — el mismo catálogo CUCoP+ (§3.4) ya cargado en memoria para resolver `clave` (§5.2), reutilizado aquí. Cuando el `cve_cucop` del ítem no está en el catálogo local (puede ser más nuevo que el snapshot descargado), no hay mejor fuente disponible y el texto degenerado queda tal cual.

### 5.3.3 Limpieza de numeración/viñetas al inicio de `producto`

Algunas instituciones anteponen a "Descripción detallada" numeración de partida, viñetas o una clave CSG con separador distinto al esperado (espacios en vez de puntos, o sin el cero inicial) — texto ajeno al nombre del medicamento. Ejemplos reales vistos en el dataset: `"13040096 UPADACITINIB..."`, `"2. 1 TEOFILINA..."`, `"-OLANZAPINA..."`, `"•\tASPIRINA..."`, `"10.   010.000.2739.00 - DIETA..."`.

`buildRegistro` limpia esto con `LEADING_JUNK_RE`: recorta cualquier prefijo de dígitos/espacios/puntos/guiones/viñetas/comillas, pero solo si justo después empieza una letra — así no se come dígitos que sí son parte del texto (`"5% DIOXIDO DE CARBONO..."`: al `5` le sigue `%`, no una letra, así que se conserva intacto). Una guarda adicional evita truncar a la mitad un sufijo de clave alfanumérico como `"...1757.P0 MELFALAN..."` (se deja el texto completo en vez de cortar a `"P0 MELFALAN..."`). Se aplica después del reemplazo por CUCoP+ de §5.3.2, así que corre sobre cualquiera de las dos fuentes por igual.

Corrida del 2026-08-10: 965 de 16,911 registros tenían este prefijo; los 3 casos protegidos por las guardas de arriba quedaron sin tocar (2 con `%` pegado al dígito inicial, 1 con sufijo de clave alfanumérico). Como efecto colateral, el agrupamiento canónico de `scripts/validar-claves.js` (agrupa por nombre+dosis extraídos de `producto`) mejoró para estos registros — cobertura de `clave` subió de 72.5% a 74.9%.

### 5.4 Reporte de calidad

Cada corrida de `extract.js` genera `docs/data.calidad.json`: top 20 valores más altos (por `valor_maximo`, para detectar outliers), registros con precio o cantidad en cero/negativo, y el detalle de cada corrección de `precio_unitario`.

### 5.5 Integración de volumen y precio entre contratos "Abierto" y "Cerrado"

Los dos tipos de contrato representan la cantidad de forma distinta (§4): "Cerrado" trae un solo número comprometido, "Abierto" trae un rango mín/máx. Mezclarlos sin criterio en una suma agregada mezclaría un "monto real comprometido" (Cerrado) con un "techo de exposición posible" (Abierto) como si fueran la misma clase de número.

La integración se resuelve con `cantidad_minima`/`cantidad_maxima` y `valor_minimo`/`valor_maximo`, siempre poblados (salvo el caso de servicios/medicina magistral sin ninguna cantidad, §5.3.1): en "Cerrado" cada par cae al mismo número (no hay rango que representar); en "Abierto" divergen. No existe un campo `cantidad`/`valor` suelto de "número único" -- se eliminó por ser redundante (siempre igual a `cantidad_maxima`/`valor_maximo`) y por invitar a sumarlo entre registros sin distinguir piso de techo. **Cualquier suma o promedio agregado debe usar explícitamente `valor_minimo` (gasto garantizado) o `valor_maximo` (techo de exposición)**, nunca mezclar ambos. `tipo_contrato` permite segmentar/filtrar antes de agregar, cuando se necesita analizar cada esquema por separado en vez de combinarlos.

---

## 6. Despliegue

**Arquitectura: un solo servicio (GitHub), sin base de datos ni backend separados.**

| Pieza | Herramienta | Rol |
|---|---|---|
| Cómputo programado | GitHub Actions (cron) | Corre el pipeline de Playwright periódicamente, en contenedor efímero |
| Almacenamiento | Archivo JSON/Excel commiteado al repo | Resultado del pipeline como dato versionado — sin motor de base de datos |
| Dashboard | GitHub Pages | HTML/JS estático que carga el archivo de datos y filtra/busca del lado del cliente |

Todo el filtrado del dashboard ocurre en JavaScript en el navegador del usuario — no hay motor de consultas en servidor. Es una decisión deliberada dado el volumen de datos (decenas de miles de filas, no millones) y el uso esperado (consulta, no transaccional).

La tabla principal muestra solo las columnas de consulta más frecuente (código y número de contrato, institución, proveedor, clave CUCoP, producto, grupo terapéutico); el resto de los campos del modelo (§4) — clave del Compendio, precio, cantidad, valores, fechas — se consulta por registro en el modal "Detalle".

El modal muestra cantidad y valor de forma condicional: si el contrato tiene rango genuino (`cantidad_minima != cantidad_maxima`, ver §5.5), muestra los 4 campos por separado (cantidad mínima/máxima, valor mínimo/máximo); si no, solo un par Cantidad/Valor (tomado de `cantidad_maxima`/`valor_maximo`, idénticos a `cantidad_minima`/`valor_minimo` en ese caso). Validado contra el dataset completo que ambos pares coinciden en si hay rango o no, salvo un caso trivial (`precio_unitario` en 0, donde el rango de valor colapsa a 0–0 aunque la cantidad sí tenga rango real) — no hay casos donde el valor tenga rango pero la cantidad no.

**Frecuencia de actualización:** pendiente de definir — el cron de GitHub Actions queda configurado pero ajustable (candidatos: semanal, diaria, manual).

**Prerrequisito:** una cuenta de GitHub (gratuita) para alojar el repositorio, correr las Actions y servir Pages.

---

## 7. Archivos del proyecto

| Archivo | Función |
|---|---|
| `scripts/build-compendio.js` | Convierte el Compendio Nacional de Medicamentos (.xlsm del CSG) a `data/compendio_medicamentos.json` |
| `scripts/build-cucop.js` | Convierte el catálogo CUCoP+ (.xlsx) a `data/cucop_medicamentos.json` — ver §3.4 |
| `scripts/download-csv.js` | Descarga el CSV anual de "Contratos de la Plataforma Integral" desde Datos Abiertos |
| `scripts/lib/dataset.js` | Piezas compartidas entre `extract.js` y `validar-claves.js` (regex de clave, diccionario de columnas, escritura de `data.xlsx`) — un solo lugar para no duplicar lo que ambos necesitan |
| `scripts/extract.js` | Pipeline principal — ver §5.1 |
| `scripts/validar-claves.js` | Sanity check de claves post-corrida — ver §5.2 |
| `docs/index.html` | Dashboard estático: búsqueda, filtros, orden por columna, paginación, modal de detalle por registro — ver §6 |
| `docs/data.json` | Dataset consolidado, un registro por ítem de contrato |
| `docs/data.xlsx` | Copia del dataset en Excel, con hoja "Diccionario" (descripción de cada columna) además de la hoja "Precios" |
| `docs/data.errores.json` | Contratos que fallaron la extracción (reintentables hasta 3 veces) |
| `docs/data.calidad.json` | Reporte de anomalías — ver §5.4 |
| `docs/data.correcciones.json` | Registro de correcciones de clave aplicadas y casos sin resolver — ver §5.2 |
| `.github/workflows/scrape.yml` | Workflow de GitHub Actions |
