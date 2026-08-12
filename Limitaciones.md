# Limitaciones de los Datos

**Aplica a:** dataset de 2026 (`docs/data.json`, 16,911 registros) — las cifras citadas abajo son de la corrida más reciente y se actualizan con cada corrida completa; ver `docs/data.calidad.json` para los números vigentes.

Este documento resume dónde puede haber datos erróneos, imprecisos o incompletos en esta base, y qué alternativas existen para cada caso. Cómo funciona la herramienta está en `Metodologia.md`; aquí solo se documenta qué puede estar mal y qué hacer al respecto.

**Idea central:** esta es una base construida por scraping de fuentes públicas, cruzada contra un catálogo oficial descargado una sola vez. Es útil para análisis de tendencias y comparación de precios, pero **no es una fuente autoritativa verificada campo por campo** — antes de usar un registro individual para una decisión importante, vale la pena revisar las limitaciones de abajo.

---

## 1. Grupo terapéutico: falta en 29.2% de los registros (4,939 de 16,911)

**Por qué pasa:** para asignar `grupo_terapeutico` hace falta la `clave` oficial del Compendio Nacional (CNIS), y que esa clave exista en el archivo del Compendio descargado. Dos causas:

- El catálogo CUCoP+ referencia **2,600 claves** bajo la partida de medicamentos, pero el Compendio Nacional descargado solo tiene **1,895** — cubre ~70% del universo de claves que se compran en la práctica. (El catálogo tiene 3,386 entradas en total bajo esa partida, pero 786 no traen ninguna clave CSG asociada en su propia ficha — ver Metodologia.md §5.3.2.)
- Buena parte de estos registros probablemente corresponden a medicamentos comprados **fuera del cuadro básico oficial** (ver punto 3), que por definición no tienen clave CNIS.

**Alternativas:**
- Conseguir una versión más reciente/completa del Compendio directamente del CSG.
- Aceptar el límite: no todo lo que se compra bajo la partida de medicamentos tiene equivalente en el cuadro básico oficial, así que 100% de cobertura puede no ser alcanzable con este enfoque.

---

## 2. Claves que pueden estar mal asignadas (no solo ausentes)

**Por qué pasa:** cuando la institución no cita la clave directamente en la descripción del contrato, el pipeline la recupera vía el código `cve_cucop` del catálogo CUCoP. Esa recuperación puede fallar cuando la institución usó un código genérico o de la dosis/presentación equivocada — confirmado contra la respuesta cruda de la API (no es un error de la extracción, viene así desde el origen).

**Cobertura de las dos capas de verificación en producción** (ver Metodologia.md §5.2):
- De los **6,015 productos canónicos** del dataset, **2,326 quedan sin resolver** tras el sanity check completo (sin clave de la descripción, sin CUCoP válido, sin match en el Compendio local) — documentados en `docs/data.correcciones.json`, sin inventar una clave.
- La autoverificación de CUCoP confía por defecto cuando el contrato o la propia ficha CUCoP+ no traen ningún número de dosis que comparar (ver Metodologia.md §5.2) — evita descartar claves correctas solo porque el catálogo no describe la dosis (ej. `TOCILIZUMAB`, `IVERMECTINA`, vacunas), sin dejar de rechazar los casos donde sí hay números y no coinciden (evidencia real de `cve_cucop` mal citado).

**Números de dosis mal formateados en el origen** (no corregido automáticamente): algunas descripciones traen la dosis partida por un espacio de más — ej. `"EMICIZUMAB 1 50 MG"` en vez de `"EMICIZUMAB 150 MG"`, confirmado que así viene desde el origen (no es un artefacto de la limpieza de prefijos del pipeline). Esto rompe cualquier comparación numérica (autoverificación de CUCoP, agrupamiento canónico, búsqueda en el Compendio local) aunque el resto del dato esté bien. No se intenta recomponer el número automáticamente — no hay forma confiable de distinguir "150" partido en "1"+"50" de dos números genuinamente separados (ej. "2 100 MG" sí puede significar "2 piezas de 100 mg"), y adivinar mal introduciría errores peores que dejar la clave sin resolver.

**Alternativas para lo sin resolver:**
- Buscar por nombre en `vademecum.es/cnis` (espejo navegable del Compendio, con más claves que nuestro archivo — 2,600 vs. 1,895) — evaluado pero no automatizado, ver punto 3.
- Descargar y parsear `https://www.csg.gob.mx/Comp26042025.pdf` (Compendio en PDF, +10MB) como fuente alterna si vademecum dejara de existir — no implementado.
- **Corrección manual dirigida** (priorizar por `valor_maximo` del contrato en vez de cubrir el 100%): tomar el `producto` de `docs/data.correcciones.json` (`sin_resolver`), determinar la clave correcta a mano, y editar `docs/data.json` poniendo `clave`, `grupo_terapeutico` y **`clave_fuente: "manual"`** en todos los registros con ese `producto`. Cualquier `clave_fuente` distinto de `null`/`"cucop"` se trata como confiable, así que la corrección sobrevive a la siguiente corrida de `validar-claves.js` y se propaga sola a otros registros del mismo producto canónico. Correr `node scripts/validar-claves.js` una vez después de editar para regenerar la propagación y el Excel.

---

## 3. Buena parte de los registros probablemente son compras fuera del cuadro básico (CNIS)

**Por qué pasa:** al probar `vademecum.es/cnis` como fuente de búsqueda por nombre, una muestra representativa de 50 productos con dosis especificada tuvo solo 2% de coincidencias, y **86% no devolvió ninguna ficha CNIS en absoluto** — ni siquiera esa fuente más completa los tiene catalogados. Esto sugiere que esos medicamentos se compraron por excepción o necesidad especializada, fuera de la lista oficial de abasto — algo permitido y real, no un error de datos.

**Por qué no se automatizó vademecum.es:** al ritmo medido (~2.5s por consulta) y ese porcentaje de acierto, cubrir los productos pendientes tomaría horas de scraping contra un sitio de terceros para un rendimiento marginal. El mecanismo (búsqueda por nombre vía `/buscar?q=...&cc=mx`, lectura de fichas `/ficha-cnis/...`) está probado y funciona — queda como alternativa manual para investigar productos específicos de alto valor, no como parte de la corrida regular.

**Alternativas:**
- Aceptar que un `grupo_terapeutico` vacío puede significar "no está en el cuadro básico" tanto como "no se pudo determinar" — hoy no se distingue entre ambos casos en el dato.
- Si se necesita esa distinción, agregar un campo adicional (ej. `en_cuadro_basico: true/false/desconocido`).

---

## 4. Contratos "por monto": `cantidad_minima`/`cantidad_maxima` se derivan, no vienen directas de la fuente

**Por qué pasa:** algunos ítems (`tipo_contrato_abierto: "MONTO"` en la API de origen) comprometen un techo de gasto en pesos, no una cantidad de piezas — la fuente no publica ninguna cantidad para estos casos, ni siquiera `precio_unitario` en algunos (servicios, medicina magistral; ahí `cantidad_minima`/`cantidad_maxima` quedan en `null` — el único caso donde eso ocurre, ver Metodologia.md §5.3.1). Sí publica `subtotal` (el "Monto de la Oferta" visible en el detalle del contrato en el sitio de Compras MX) — ver Metodologia.md §5.3.1 para cómo se usa para derivar `cantidad_minima`/`cantidad_maxima` (ambas iguales, sin rango) y respaldar `valor_minimo`/`valor_maximo`.

**Riesgo residual:** la cantidad derivada asume que `subtotal` es exactamente `precio_unitario × cantidad` sin redondeos raros del proveedor; no se ha visto un caso donde no cuadre, pero no está garantizado matemáticamente.

**Alternativas:** revisar `docs/data.calidad.json` (`cantidades_derivadas_de_subtotal_entre_precio_unitario`) para ver todos los casos donde se aplicó esta derivación.

---

## 5. `cantidad_minima`/`cantidad_maxima` son el compromiso contractual, no lo entregado realmente

**Por qué pasa:** las fuentes de contratación pública solo exponen la cantidad pactada en el contrato. No existe una fuente pública que reporte cuánto se entregó/facturó realmente — puede haber diferencia si el contrato no se ejecutó completo.

**Alternativas:** ninguna fuente pública cubre esto hoy. Una solicitud de transparencia (INAI) por datos de ejecución real es posible, pero fuera del alcance de un pipeline automatizado.

---

## 6. `precio_unitario` puede venir corrupto desde el origen

**Por qué pasa:** el campo de la API puede traer copiado por error el valor de otro campo (ej. el monto total de una compra grande en vez del precio unitario) — un error del propio sistema de Compras MX, no de la extracción.

**Qué se hizo:** el pipeline recalcula `precio_unitario = subtotal / cantidad_minima` y solo usa el valor de la API cuando coincide (±1%). Cada corrección queda registrada en `docs/data.calidad.json`.

**Riesgo residual:** si `subtotal` también viniera corrupto en algún caso no detectado, el precio recalculado heredaría el error.

**Alternativas:** revisar `docs/data.calidad.json` (top-20 valores más altos) antes de usar el dataset para análisis de montos totales — la forma más rápida de detectar outliers residuales.

---

## 7. Texto de producto: cuatro patrones degenerados corregidos automáticamente; typos sueltos no

**Qué se detectó, verificado en vivo contra el sitio:** la institución compradora a veces captura la "Descripción detallada" del ítem de forma degenerada, en cuatro patrones estructurales recurrentes:
- Referencia vacía de contenido: el texto es literalmente `CONFORME A PARTIDA N DE LA CONVOCATORIA` — remite a su propia partida sin describir el producto (128 registros en la corrida donde se detectó).
- Sin espacios entre palabras: el texto sí describe el producto pero viene concatenado, ej. `ACICLOVIR200MGENVASECON25COMPRIMIDOSOTABLETAS...` (134 registros).
- Numeración de partida, viñetas o una clave con separador raro pegadas al inicio, ajenas al nombre del medicamento: ej. `13040096 UPADACITINIB...`, `2. 1 TEOFILINA...`, `-OLANZAPINA...`, `•\tASPIRINA...` (965 registros).
- La misma coletilla `CONFORME A PARTIDA N DE LA CONVOCATORIA` del primer patrón, pero pegada al **final** de una descripción real y completa (no en vez de ella): ej. `PARACETAMOL 500 MG ENVASE CON 10 TABLETAS.CONFORME A PARTIDA 204 DE LA CONVOCATORIA` (612 registros — población disjunta del primer patrón: 0 registros coinciden con ambos a la vez, confirmado contra el dataset).

**Qué se hizo:** para el primer y segundo patrón, el pipeline sustituye `producto` por la ficha del catálogo CUCoP+, que siempre viene bien formada (Metodologia.md §5.3.2). El tercero se recorta con una regex que exige que el prefijo termine justo antes de una letra, para no comerse dígitos que sí son parte del texto (ej. `5% DIOXIDO DE CARBONO...` se conserva intacto) — Metodologia.md §5.3.3. El cuarto recorta solo la coletilla del final, conservando la descripción real — Metodologia.md §5.3.4. Cobertura completa en la corrida vigente — 0 casos de los patrones 1, 2 y 4, y solo 3 del tercero (protegidos a propósito por las guardas de la regex: 2 con `%` pegado al dígito inicial, 1 con sufijo de clave alfanumérico truncable) en los 16,911 registros.

**Lo que sigue sin corregir:** errores de digitación sueltos (typos evidentes que no siguen ninguno de los cuatro patrones de arriba) no se detectan ni corrigen.

**Alternativas:**
- No hay forma automática de "arreglar" texto corrupto de origen sin arriesgar inventar datos.
- Se podría ampliar la heurística a otros patrones estructurales si aparecen como recurrentes (ej. flag si el texto no contiene ningún nombre reconocible de principio activo). No implementado.
- Para typos sueltos, la única vía es corrección manual dirigida editando `producto` directo en `docs/data.json` — a diferencia de `clave` (punto 2), no hay un campo `producto_fuente` que la marque como confiable, así que no sobrevive a la siguiente corrida completa a menos que también se corrija en el origen.

---

## 8. Un medicamento puede tener más de un grupo terapéutico

**Por qué pasa:** el Compendio CSG asigna a veces más de un grupo terapéutico al mismo medicamento. `grupo_terapeutico` se guarda como arreglo (puede traer varios valores), sin regla de desempate a un único valor.

**Alternativas:** si un análisis necesita un solo grupo por registro, definir una regla de prioridad (ej. el primero listado, o el más específico).

---

## 9. Fragilidad de la extracción (dependencia de automatización de navegador)

**Por qué pasa:** Compras MX no tiene API pública — todo se extrae automatizando un navegador real contra una SPA protegida con reCAPTCHA v3. Un cambio futuro en el sitio (nueva versión de la SPA, cambio de estructura, nuevo mecanismo de paginación) podría romper la extracción sin aviso.

**Alternativas:** ninguna mientras Compras MX no publique una API — es el costo aceptado de esta arquitectura. Mitigación parcial: el reporte de errores por corrida (`docs/data.errores.json`) sirve de alerta temprana si algo empieza a fallar sistemáticamente. La búsqueda del contrato dentro de la tabla paginada (`findAndClickContrato`) tiene su propio límite de tiempo interno (no solo el timeout duro externo, que no cancela trabajo en curso en JS) para nunca quedarse colgada indefinidamente si el contrato no aparece en la tabla.

---

## 10. Cobertura histórica limitada a 2020–2026

**Por qué pasa:** los archivos de "Contratos de la Plataforma Integral" cubren de forma consistente 2020–2026. Años anteriores (CompraNet 5.0 y 3.0) usan esquemas distintos y más limitados.

**Alternativas:** extender el pipeline a esas fuentes es técnicamente posible pero requiere adaptar el parser a un esquema diferente.

---

## 11. Supuestos no verificados activamente

- **Moneda:** el pipeline asume MXN sin verificar contra la columna "Moneda" del CSV. Hoy da 100% MXN en los datos filtrados, pero no hay alerta si eso cambiara.
- **Filas descartadas por campos faltantes** (sin "Dirección del anuncio" o "Código del contrato"): no se cuentan ni se loguean. Hoy da cero filas afectadas, pero no hay visibilidad si eso cambiara.

**Alternativas:** agregar validación explícita y logging para ambos casos — no implementado.

---

## 12. Errores de extracción residuales

**Estado actual:** un puñado de contratos con error de extracción por corrida (timeouts de red), dentro del presupuesto de 3 reintentos automáticos — se resuelven solos en la siguiente corrida. Ver `docs/data.errores.json` para el detalle actualizado.

**Alternativas:** ninguna acción necesaria mientras el conteo se mantenga bajo — un salto grande en una corrida futura es señal de que algo cambió en el sitio fuente (ver punto 9). Si un contrato agota los 3 reintentos automáticos (queda como "fallo permanente", ya no se reintenta solo), se puede forzar manualmente: ubicarlo por `codigoContrato` en `data/raw/contratos_{año}.csv`, revisar el expediente directo en `https://comprasmx.buengobierno.gob.mx/sitiopublico/#/sitiopublico/detalle/{hash}/procedimiento` para diagnosticar, y borrar sus entradas de `docs/data.errores.json` antes de correr `extract.js` de nuevo para que se reintente.

---

## 13. `institución` puede no ser la institución real en compras consolidadas

**Por qué pasa:** `institución` se toma de "Siglas de la Institución" en el CSV. En procedimientos consolidados, esa columna puede no corresponder a la institución que realmente contrató cada renglón — la institución real solo aparece en la tabla de contratos de la página de detalle, no en el CSV (verificado con un caso real). En procedimientos NO consolidados (una sola institución, sin ambigüedad) esto no debería pasar.

**Alternativas:**
- Filtrar/interpretar con cautela cuando `tipo_procedimiento = CONSOLIDADA` — la institución ahí es menos confiable que en NO CONSOLIDADA.
- Scrapear "Dependencia o entidad contratante" de la tabla de contratos de la página de detalle (ya se visita esa página) para los casos consolidados — evaluado, no implementado por decisión explícita de mantener el pipeline simple.

---

## 14. `fecha_fallo` ausente en algunas compras consolidadas

**Por qué pasa:** el campo ya viene vacío así en el CSV fuente — el pipeline solo lo copia, no lo transforma. Afecta **856 de 16,920 registros (641 contratos únicos)**, y **el 100% son compras consolidadas** (`tipo_procedimiento = CONSOLIDADA`). Es todo-o-nada por procedimiento: de 1,138 procedimientos con partida 25301, **150 (13.2%)** carecen de `fecha_fallo` en *todos* sus contratos, y los 988 restantes la tienen en *todos* — no se encontró ni un solo procedimiento mixto (algunos contratos con fallo, otros sin él). Explicación probable: "Fecha de fallo" registra el acto de adjudicación del procedimiento completo, y en compras consolidadas (varias instituciones bajo un mismo procedimiento) ese dato a veces no se replica hacia el registro exportado en el CSV masivo — un hueco de publicación del sistema fuente para ciertos procedimientos, no un error de captura del pipeline.

**Alternativas:**
- Ninguna automatizable: el dato no está disponible en el CSV masivo para estos procedimientos. Se podría intentar buscarlo en la página de detalle del expediente (`https://comprasmx.buengobierno.gob.mx/sitiopublico/#/sitiopublico/detalle/{hash}/procedimiento`, ya visitada durante la extracción) para los 150 procedimientos afectados — no implementado, mismo criterio de simplicidad que el punto 13.
- Aceptar el vacío: `fecha_firma_contrato` (siempre presente) sigue siendo confiable para ordenar/filtrar por fecha aunque falte `fecha_fallo`.

---

## 15. Un registro trae `cantidad_minima` > `cantidad_maxima`

**Por qué pasa:** viene así invertido desde la propia API de Compras MX (`C-2026-00068873`, contrato "Cerrado" con min=200 y max=20) — no es un bug del pipeline, que preserva el rango tal cual lo reporta la fuente cuando ambos valores vienen poblados y distintos.

**Alternativas:** ninguna sin volver a la fuente original para confirmar cuál de los dos números es el correcto. Caso aislado (1 de 16,920 registros); revisar `docs/data.json` si se necesita excluirlo puntualmente.
