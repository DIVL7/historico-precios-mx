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

### 3.4 Catálogo CUCoP+

Fuente: `data/raw/cucop.xlsx`. Para toda entrada de la partida 25301, la columna `DESCRIPCIÓN` trae la clave del Compendio como prefijo — es estático (no depende de qué institución compre), así que sirve de respaldo cuando la descripción libre del contrato no incluye la clave directamente.

---

## 4. Modelo de datos

| Campo | Fuente | Notas |
|---|---|---|
| `codigo_contrato` | Compras MX | Identificador único a nivel contrato |
| `clave` | Ver §5.2 (asignación y validación de clave) | Formato CSG `NNN.NNN.NNNN.NN`, o `null` si no se pudo determinar |
| `clave_fuente` | Interno | De dónde salió `clave`: `descripcion` (del texto del contrato), `cucop` (vía catálogo CUCoP), `validacion_nombre_local` (por nombre contra el Compendio), `propagacion_confiable` (heredada de otro registro del mismo producto), o `null` |
| `producto` | "Descripción detallada" (`detallepartidas`), texto restante tras remover la clave | |
| `tipo_insumo` | Constante `MEDICAMENTO` | |
| `grupo_terapéutico` | Join contra Compendio CSG por `clave` | `null` si `clave` es `null` o no existe en el Compendio |
| `procedimiento` | "Número de procedimiento" | |
| `tipo_procedimiento` | "Compra consolidada" SI/NO → CONSOLIDADA / NO CONSOLIDADA | Consolidada: varias instituciones agrupan sus necesidades en un solo procedimiento (coordinado por una entidad líder) en vez de licitar cada una por separado. No consolidada: una institución licita y contrata por su cuenta |
| `proveedor` | "Proveedor o contratista" | |
| `institución` | "Siglas de la Institución" (CSV) | En compras consolidadas puede no corresponder a la institución que realmente contrató cada renglón — ver `Limitaciones.md` |
| `precio_unitario` | "Precio unitario sin impuestos" (`detallepartidas`), recalculado si es inconsistente con el subtotal — ver §5.3 | |
| `cantidad` | `cantidad_maxima ?? cantidad ?? cantidad_minima`, o derivada de `subtotal / precio_unitario` si las tres vienen vacías — ver §5.3.1 | Usado para `valor`. Es el compromiso contractual, no el volumen entregado |
| `cantidad_minima` / `cantidad_maxima` | "Cantidad mínima" / "Cantidad máxima" (`detallepartidas`) | `cantidad_maxima` solo viene poblada en contratos tipo "abierto" |
| `tipo_contrato` | "Tipo de contrato" (CSV) | `Abierto` / `Cerrado`, valor oficial de la fuente — no se infiere de si `cantidad_maxima` viene poblada, para tener una sola fuente de verdad sin ambigüedad al integrar ambos esquemas de cantidad en la misma tabla |
| `valor` | `precio_unitario × cantidad`, o `subtotal` directo si no se puede calcular así — ver §5.3.1 | Número único de referencia para ver un registro individual — ver §5.5 antes de sumarlo entre muchos registros |
| `valor_minimo` / `valor_maximo` | `precio_unitario × cantidad_minima` / `precio_unitario × cantidad_maxima`, o `subtotal` directo si no se puede calcular así — ver §5.3.1 | Piso y techo de exposición contractual — ver §5.5 |
| `fecha_firma_contrato` | "Fecha firma contrato" (CSV) | Cuándo se formalizó el contrato |
| `fecha_fallo` | "Fecha de fallo" (CSV) | Cuándo se determinó el precio ganador (adjudicación) — normalmente precede a la firma |
| `fecha_inicio_contrato` / `fecha_fin_contrato` | "Fecha de inicio del contrato" / "Fecha de fin del contrato" (CSV) | Ventana de vigencia del contrato — cuándo ese precio estuvo activo para entregas, distinto de cuándo se firmó |

---

## 5. Pipeline

### 5.1 Extracción (`scripts/extract.js`)

1. Descargar el CSV anual desde Datos Abiertos (`scripts/download-csv.js`, vía navegador automatizado).
2. Filtrar filas donde `Partida específica` contiene `25301`.
3. Agrupar filas por expediente (hash extraído de `Dirección del anuncio`).
4. Por cada expediente, navegar al detalle y localizar cada contrato dentro de la tabla paginada de resultados.
5. Extraer `detallepartidas/{hash}/{codigo_contrato}` — la respuesta JSON trae los ítems ya estructurados (`cve_cucop`, `descripcion`, `precio_unitario`, `subtotal`, `cantidad`, `cantidad_minima`, `cantidad_maxima`, `um`).
6. Por cada ítem, asignar clave (§5.2), calcular `precio_unitario` corregido (§5.3) y `valor`, y construir el registro final.
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

### 5.3.1 Derivación de `cantidad` y `valor` en contratos "por monto"

Un subconjunto de ítems trae `tipo_contrato_abierto: "MONTO"` en la respuesta de `detallepartidas`: el compromiso contractual es un techo de gasto, no una cantidad de piezas, así que `cantidad`, `cantidad_minima` y `cantidad_maxima` vienen vacíos desde el origen (no es un hueco de la extracción). Sí trae `subtotal` — el "Monto de la Oferta" visible en el detalle del contrato en el sitio — que permite derivar dos cosas:

- **`cantidad`** (informativa): `subtotal / precio_unitario`, redondeada al entero más cercano, solo cuando las tres vías directas fallan. Cada caso queda registrado en `docs/data.calidad.json` (`cantidades_derivadas_de_subtotal_entre_precio_unitario`).
- **`valor` / `valor_minimo` / `valor_maximo`**: se toman de `subtotal` directo, no de `precio_unitario × cantidad_derivada` — multiplicar de vuelta sería un viaje de ida y vuelta innecesario que puede perder precisión por el redondeo intermedio de `cantidad`. Este respaldo directo aplica siempre que no haya un rango `cantidad_minima`/`cantidad_maxima` genuino que preservar (ver §5.5), incluyendo el caso donde `precio_unitario` viene en `0` desde el origen (visto en servicios y medicina magistral): ahí no se puede derivar `cantidad` tampoco, pero `valor` sí se rescata de `subtotal`.

### 5.4 Reporte de calidad

Cada corrida de `extract.js` genera `docs/data.calidad.json`: top 20 valores más altos (para detectar outliers), registros con precio o cantidad en cero/negativo, y el detalle de cada corrección de `precio_unitario`.

### 5.5 Integración de volumen y precio entre contratos "Abierto" y "Cerrado"

Los dos tipos de contrato representan la cantidad de forma distinta (§4): "Cerrado" trae un solo número comprometido, "Abierto" trae un rango mín/máx. Mezclarlos sin criterio en una suma agregada mezclaría un "monto real comprometido" (Cerrado) con un "techo de exposición posible" (Abierto) como si fueran la misma clase de número.

La integración se resuelve con tres campos, cada uno con un propósito distinto:

- **`valor`** — número único (`precio_unitario × cantidad`, con `cantidad` = techo en Abierto). Sirve para ver un registro individual, no para sumar entre muchos.
- **`valor_minimo` / `valor_maximo`** — piso y techo de exposición contractual, bien definidos para ambos tipos: en "Cerrado" son iguales entre sí (no hay rango que representar); en "Abierto" divergen. **Cualquier suma o promedio agregado debe usar uno de estos dos, no `valor`** — `valor_minimo` para el gasto garantizado, `valor_maximo` para el techo de exposición.
- **`tipo_contrato`** — permite segmentar/filtrar antes de agregar, cuando se necesita analizar cada esquema por separado en vez de combinarlos.

---

## 6. Despliegue

**Arquitectura: un solo servicio (GitHub), sin base de datos ni backend separados.**

| Pieza | Herramienta | Rol |
|---|---|---|
| Cómputo programado | GitHub Actions (cron) | Corre el pipeline de Playwright periódicamente, en contenedor efímero |
| Almacenamiento | Archivo JSON/Excel commiteado al repo | Resultado del pipeline como dato versionado — sin motor de base de datos |
| Dashboard | GitHub Pages | HTML/JS estático que carga el archivo de datos y filtra/busca del lado del cliente |

Todo el filtrado del dashboard ocurre en JavaScript en el navegador del usuario — no hay motor de consultas en servidor. Es una decisión deliberada dado el volumen de datos (decenas de miles de filas, no millones) y el uso esperado (consulta, no transaccional).

**Frecuencia de actualización:** pendiente de definir — el cron de GitHub Actions queda configurado pero ajustable (candidatos: semanal, diaria, manual).

**Prerrequisito:** una cuenta de GitHub (gratuita) para alojar el repositorio, correr las Actions y servir Pages.

---

## 7. Archivos del proyecto

| Archivo | Función |
|---|---|
| `scripts/build-compendio.js` | Convierte el Compendio Nacional de Medicamentos (.xlsm del CSG) a `data/compendio_medicamentos.json` |
| `scripts/download-csv.js` | Descarga el CSV anual de "Contratos de la Plataforma Integral" desde Datos Abiertos |
| `scripts/extract.js` | Pipeline principal — ver §5.1 |
| `scripts/validar-claves.js` | Sanity check de claves post-corrida — ver §5.2 |
| `docs/index.html` | Dashboard estático: búsqueda, filtros, orden por columna, paginación |
| `docs/data.json` / `docs/data.xlsx` | Dataset consolidado |
| `docs/data.errores.json` | Contratos que fallaron la extracción (reintentables hasta 3 veces) |
| `docs/data.calidad.json` | Reporte de anomalías — ver §5.4 |
| `docs/data.correcciones.json` | Registro de correcciones de clave aplicadas y casos sin resolver — ver §5.2 |
| `.github/workflows/scrape.yml` | Workflow de GitHub Actions |
