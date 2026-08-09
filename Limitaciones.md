# Limitaciones de los Datos

**Aplica a:** dataset de 2026 (`docs/data.json`, 24,875 registros)

Este documento resume dónde puede haber datos erróneos, imprecisos o incompletos en esta base, y qué alternativas existen para cada caso. Cómo funciona la herramienta está en `Metodologia.md`; aquí solo se documenta qué puede estar mal y qué hacer al respecto.

**Idea central:** esta es una base construida por scraping de fuentes públicas, cruzada contra un catálogo oficial descargado una sola vez. Es útil para análisis de tendencias y comparación de precios, pero **no es una fuente autoritativa verificada campo por campo** — antes de usar un registro individual para una decisión importante, vale la pena revisar las limitaciones de abajo.

---

## 1. Grupo terapéutico: falta en 48.3% de los registros (12,020 de 24,875)

**Por qué pasa:** para asignar `grupo_terapeutico` hace falta la `clave` oficial del Compendio Nacional (CNIS), y que esa clave exista en el archivo del Compendio descargado. Dos causas:

- El catálogo CUCoP+ referencia **2,596 claves** bajo la partida de medicamentos, pero el Compendio Nacional descargado solo tiene **1,895** — cubre ~70% del universo de claves que se compran en la práctica.
- Buena parte de estos registros probablemente corresponden a medicamentos comprados **fuera del cuadro básico oficial** (ver punto 3), que por definición no tienen clave CNIS.

**Alternativas:**
- Conseguir una versión más reciente/completa del Compendio directamente del CSG.
- Aceptar el límite: no todo lo que se compra bajo la partida de medicamentos tiene equivalente en el cuadro básico oficial, así que 100% de cobertura puede no ser alcanzable con este enfoque.

---

## 2. Claves que pueden estar mal asignadas (no solo ausentes)

**Por qué pasa:** cuando la institución no cita la clave directamente en la descripción del contrato, el pipeline la recupera vía el código `cve_cucop` del catálogo CUCoP. Esa recuperación puede fallar cuando la institución usó un código genérico o de la dosis/presentación equivocada — confirmado contra la respuesta cruda de la API (no es un error de la extracción, viene así desde el origen).

**Cobertura de las dos capas de verificación en producción** (ver Metodologia.md §5.2):
- De los registros con clave vía CUCoP, **33% no pasan** la autoverificación contra la dosis que ese código declara.
- Del total de productos únicos del dataset, **7,302 quedan sin resolver** tras el sanity check completo (sin clave de la descripción, sin CUCoP válido, sin match en el Compendio local) — documentados en `docs/data.correcciones.json`, sin inventar una clave.

**Alternativas para lo sin resolver:**
- Buscar por nombre en `vademecum.es/cnis` (espejo navegable del Compendio, con más claves que nuestro archivo — 2,596 vs. 1,895) — evaluado pero no automatizado, ver punto 3.
- Descargar y parsear `https://www.csg.gob.mx/Comp26042025.pdf` (Compendio en PDF, +10MB) como fuente alterna si vademecum dejara de existir — no implementado.
- **Corrección manual dirigida** (priorizar por `valor` del contrato en vez de cubrir el 100%): tomar el `producto` de `docs/data.correcciones.json` (`sin_resolver`), determinar la clave correcta a mano, y editar `docs/data.json` poniendo `clave`, `grupo_terapeutico` y **`clave_fuente: "manual"`** en todos los registros con ese `producto`. Cualquier `clave_fuente` distinto de `null`/`"cucop"` se trata como confiable, así que la corrección sobrevive a la siguiente corrida de `validar-claves.js` y se propaga sola a otros registros del mismo producto canónico. Correr `node scripts/validar-claves.js` una vez después de editar para regenerar la propagación y el Excel.

---

## 3. Buena parte de los registros probablemente son compras fuera del cuadro básico (CNIS)

**Por qué pasa:** al probar `vademecum.es/cnis` como fuente de búsqueda por nombre, una muestra representativa de 50 productos con dosis especificada tuvo solo 2% de coincidencias, y **86% no devolvió ninguna ficha CNIS en absoluto** — ni siquiera esa fuente más completa los tiene catalogados. Esto sugiere que esos medicamentos se compraron por excepción o necesidad especializada, fuera de la lista oficial de abasto — algo permitido y real, no un error de datos.

**Por qué no se automatizó vademecum.es:** al ritmo medido (~2.5s por consulta) y ese porcentaje de acierto, cubrir los productos pendientes tomaría horas de scraping contra un sitio de terceros para un rendimiento marginal. El mecanismo (búsqueda por nombre vía `/buscar?q=...&cc=mx`, lectura de fichas `/ficha-cnis/...`) está probado y funciona — queda como alternativa manual para investigar productos específicos de alto valor, no como parte de la corrida regular.

**Alternativas:**
- Aceptar que un `grupo_terapeutico` vacío puede significar "no está en el cuadro básico" tanto como "no se pudo determinar" — hoy no se distingue entre ambos casos en el dato.
- Si se necesita esa distinción, agregar un campo adicional (ej. `en_cuadro_basico: true/false/desconocido`).

---

## 4. `cantidad` es el compromiso contractual, no lo entregado realmente

**Por qué pasa:** las fuentes de contratación pública solo exponen la cantidad pactada en el contrato. No existe una fuente pública que reporte cuánto se entregó/facturó realmente — puede haber diferencia si el contrato no se ejecutó completo.

**Alternativas:** ninguna fuente pública cubre esto hoy. Una solicitud de transparencia (INAI) por datos de ejecución real es posible, pero fuera del alcance de un pipeline automatizado.

---

## 5. `precio_unitario` puede venir corrupto desde el origen

**Por qué pasa:** el campo de la API puede traer copiado por error el valor de otro campo (ej. el monto total de una compra grande en vez del precio unitario) — un error del propio sistema de Compras MX, no de la extracción.

**Qué se hizo:** el pipeline recalcula `precio_unitario = subtotal / cantidad_minima` y solo usa el valor de la API cuando coincide (±1%). Cada corrección queda registrada en `docs/data.calidad.json`.

**Riesgo residual:** si `subtotal` también viniera corrupto en algún caso no detectado, el precio recalculado heredaría el error.

**Alternativas:** revisar `docs/data.calidad.json` (top-20 valores más altos) antes de usar el dataset para análisis de montos totales — la forma más rápida de detectar outliers residuales.

---

## 6. Texto de producto corrupto o mal capturado en la fuente

**Por qué pasa:** algunas descripciones de origen tienen errores de digitación evidentes de la institución compradora, y en al menos un caso el campo `producto` capturó el título genérico del contrato en vez del nombre del medicamento.

**Alternativas:**
- No hay forma automática de "arreglar" texto corrupto de origen sin arriesgar inventar datos.
- Se podría agregar una heurística de detección (ej. flag si el texto no contiene ningún nombre reconocible de principio activo) para marcar estos casos como sospechosos, sin corregirlos automáticamente. No implementado.

---

## 7. Un medicamento puede tener más de un grupo terapéutico

**Por qué pasa:** el Compendio CSG asigna a veces más de un grupo terapéutico al mismo medicamento. `grupo_terapeutico` se guarda como arreglo (puede traer varios valores), sin regla de desempate a un único valor.

**Alternativas:** si un análisis necesita un solo grupo por registro, definir una regla de prioridad (ej. el primero listado, o el más específico).

---

## 8. Fragilidad de la extracción (dependencia de automatización de navegador)

**Por qué pasa:** Compras MX no tiene API pública — todo se extrae automatizando un navegador real contra una SPA protegida con reCAPTCHA v3. Un cambio futuro en el sitio (nueva versión de la SPA, cambio de estructura, nuevo mecanismo de paginación) podría romper la extracción sin aviso.

**Alternativas:** ninguna mientras Compras MX no publique una API — es el costo aceptado de esta arquitectura. Mitigación parcial: el reporte de errores por corrida (`docs/data.errores.json`) sirve de alerta temprana si algo empieza a fallar sistemáticamente.

---

## 9. Cobertura histórica limitada a 2020–2026

**Por qué pasa:** los archivos de "Contratos de la Plataforma Integral" cubren de forma consistente 2020–2026. Años anteriores (CompraNet 5.0 y 3.0) usan esquemas distintos y más limitados.

**Alternativas:** extender el pipeline a esas fuentes es técnicamente posible pero requiere adaptar el parser a un esquema diferente.

---

## 10. Supuestos no verificados activamente

- **Moneda:** el pipeline asume MXN sin verificar contra la columna "Moneda" del CSV. Hoy da 100% MXN en los datos filtrados, pero no hay alerta si eso cambiara.
- **Filas descartadas por campos faltantes** (sin "Dirección del anuncio" o "Código del contrato"): no se cuentan ni se loguean. Hoy da cero filas afectadas, pero no hay visibilidad si eso cambiara.

**Alternativas:** agregar validación explícita y logging para ambos casos — no implementado.

---

## 11. Errores de extracción residuales

**Estado actual:** un puñado de contratos con error de extracción por corrida (timeouts de red), dentro del presupuesto de 3 reintentos automáticos — se resuelven solos en la siguiente corrida. Ver `docs/data.errores.json` para el detalle actualizado.

**Alternativas:** ninguna acción necesaria mientras el conteo se mantenga bajo — un salto grande en una corrida futura es señal de que algo cambió en el sitio fuente (ver punto 8). Si un contrato agota los 3 reintentos automáticos (queda como "fallo permanente", ya no se reintenta solo), se puede forzar manualmente: ubicarlo por `codigoContrato` en `data/raw/contratos_{año}.csv`, revisar el expediente directo en `https://comprasmx.buengobierno.gob.mx/sitiopublico/#/sitiopublico/detalle/{hash}/procedimiento` para diagnosticar, y borrar sus entradas de `docs/data.errores.json` antes de correr `extract.js` de nuevo para que se reintente.

---

## 12. `institución` puede no ser la institución real en compras consolidadas

**Por qué pasa:** `institución` se toma de "Siglas de la Institución" en el CSV. En procedimientos consolidados, esa columna puede no corresponder a la institución que realmente contrató cada renglón — la institución real solo aparece en la tabla de contratos de la página de detalle, no en el CSV (verificado con un caso real). En procedimientos NO consolidados (una sola institución, sin ambigüedad) esto no debería pasar.

**Alternativas:**
- Filtrar/interpretar con cautela cuando `tipo_procedimiento = CONSOLIDADA` — la institución ahí es menos confiable que en NO CONSOLIDADA.
- Scrapear "Dependencia o entidad contratante" de la tabla de contratos de la página de detalle (ya se visita esa página) para los casos consolidados — evaluado, no implementado por decisión explícita de mantener el pipeline simple.
