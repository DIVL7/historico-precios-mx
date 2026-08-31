#!/bin/bash
# Uso: scripts/run-extract.sh data/raw/contratos_2025.csv [--concurrency 4]
#
# Relanza `node scripts/extract.js` automáticamente hasta que ya no queden
# expedientes pendientes, y SIEMPRE registra la salida en data/run_<origen>.log.
# Existe para no repetir dos fallas de la sesión 2026-08-10: perder horas de
# progreso porque el proceso se corta solo (ventanas de background
# inconsistentes) sin que nadie lo relance, y olvidar redirigir la salida al
# log del dataset.
#
# --reset NO se acepta acá a propósito: si un reintento automático lo
# repitiera, cada relanzamiento borraría el progreso que el intento anterior
# acaba de guardar. Para reiniciar un origen, correr una sola vez a mano:
#   node scripts/extract.js <csv> --reset --concurrency N
# y recién después usar este wrapper para el resto de la corrida.
set -uo pipefail

if [ -z "${1:-}" ]; then
  echo "Uso: scripts/run-extract.sh <archivo.csv> [flags de extract.js, sin --reset]"
  exit 1
fi
CSV="$1"; shift

for a in "$@"; do
  if [ "$a" = "--reset" ]; then
    echo "ERROR: --reset no se acepta en run-extract.sh (ver comentario en el archivo). Corré ese paso aparte, una sola vez, con node scripts/extract.js directamente."
    exit 1
  fi
  if [ "$a" = "--out" ]; then
    echo "ERROR: --out no se acepta en run-extract.sh -- al terminar la corrida este wrapper siempre valida y audita docs/data.<año>.json (rutas estándar dentro de validar-claves.js y audit-dataset.js, ver scripts/lib/dataset.js), así que un --out a otro archivo dejaría ese archivo sin sanity-check, auditoría ni Excel. Corré extract.js directo si necesitás una salida a medida."
    exit 1
  fi
done

# La regla "primer run de 4 dígitos en el nombre del archivo" duplica a propósito
# la de parseArgs() en extract.js (no hay forma limpia de pedirle el origen ya
# calculado sin resolver antes qué archivo de log usar) -- si esa regla cambia
# ahí, hay que replicarla acá.
ORIGEN=$(basename "$CSV" | grep -oE '[0-9]{4}' | head -1)
ORIGEN="${ORIGEN:-$(basename "$CSV" .csv)}"
LOG="data/run_${ORIGEN}.log"
# Antes 30 (~5min de reintentos antes de rendirse): diagnosticado en vivo el
# 2026-08-30 que el corte real es agotamiento de RAM del sistema (ver
# comentario en scripts/extract.js junto a chromium.launch()), no un
# expediente atorado -- cada intento ya es seguro (checkpoint + lock +
# fallo-permanente por contrato tras 3 tries reales en extract.js), así que
# rendirse a los 30 solo apagaba una corrida sana que iba a seguir
# progresando sola. Subido a un número alto en vez de sacar el tope del todo
# para no perder el aviso de "expediente realmente atorado" si el día que de
# verdad no avanza nada, alguien tiene que enterarse igual.
MAX_INTENTOS=1000

{ echo ""; echo "=== run-extract.sh $(date -Iseconds): $CSV $* ==="; } >> "$LOG"

for i in $(seq 1 "$MAX_INTENTOS"); do
  echo ">>> [run-extract] intento $i/$MAX_INTENTOS" | tee -a "$LOG"
  # Salida de ESTA invocación aparte del log acumulado: un `tail -N` fijo
  # sobre $LOG (que junta TODAS las corridas de este origen, ver comentario
  # de arriba) se rompe apenas una corrida imprime más de N líneas de log
  # por expediente (pasa fácil con cientos de expedientes) -- la línea "X
  # pendientes de procesar", impresa una sola vez al arrancar, queda fuera
  # de la ventana y PENDIENTES nunca matchea, aunque la corrida sí haya
  # terminado. Grepeando el archivo temporal (acotado a exactamente esta
  # invocación) no hay ventana que se pueda quedar corta.
  TMP_OUT=$(mktemp)
  node scripts/extract.js "$CSV" "$@" 2>&1 | tee -a "$LOG" "$TMP_OUT"
  PENDIENTES=$(grep -oE '[0-9]+ pendientes de procesar' "$TMP_OUT" | tail -1 | grep -oE '^[0-9]+')
  rm -f "$TMP_OUT"
  if [ "${PENDIENTES:-1}" = "0" ]; then
    echo ">>> [run-extract] sin pendientes -- corrida completa." | tee -a "$LOG"
    node scripts/validar-claves.js 2>&1 | tee -a "$LOG"
    VALIDAR_OK="${PIPESTATUS[0]}"
    node scripts/audit-dataset.js 2>&1 | tee -a "$LOG"
    AUDIT_OK="${PIPESTATUS[0]}"
    # Sin chequear esto, un crash de cualquiera de los dos (ej. docs/data.json
    # mal formado, catálogo faltante) igual terminaba en "exit 0" -- cualquier
    # automatización que confíe en el código de salida de este wrapper creería
    # que la corrida quedó validada y auditada cuando en realidad ese paso
    # falló en silencio.
    if [ "$VALIDAR_OK" != "0" ] || [ "$AUDIT_OK" != "0" ]; then
      echo ">>> [run-extract] ERROR: validar-claves.js (exit $VALIDAR_OK) o audit-dataset.js (exit $AUDIT_OK) fallaron -- ver arriba. La extracción terminó pero la post-verificación no; revisar antes de confiar en docs/data.json." | tee -a "$LOG"
      exit 1
    fi
    exit 0
  fi
  # Pausa corta antes del siguiente intento: si el intento de arriba fue
  # rechazado por el lock de extract.js (una corrida anterior sigue viva),
  # esta línea nunca imprime "N pendientes de procesar" -- sin esta pausa, el
  # loop reintentaría contra el mismo lock vivo en un ciclo apretado,
  # quemando los MAX_INTENTOS en segundos y reportando "se agotaron los
  # intentos" aunque la corrida real siga perfectamente sana.
  sleep 10
done

echo ">>> [run-extract] se agotaron $MAX_INTENTOS intentos sin terminar -- revisar manualmente (puede ser un expediente realmente atorado, no solo cortes de ventana)." | tee -a "$LOG"
exit 1
