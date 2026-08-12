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
    echo "ERROR: --out no se acepta en run-extract.sh -- al terminar la corrida este wrapper siempre valida docs/data.json (ruta fija dentro de validar-claves.js), así que un --out a otro archivo dejaría ese archivo sin sanity-check ni Excel. Corré extract.js directo si necesitás una salida a medida."
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
MAX_INTENTOS=30

{ echo ""; echo "=== run-extract.sh $(date -Iseconds): $CSV $* ==="; } >> "$LOG"

for i in $(seq 1 "$MAX_INTENTOS"); do
  echo ">>> [run-extract] intento $i/$MAX_INTENTOS" | tee -a "$LOG"
  node scripts/extract.js "$CSV" "$@" 2>&1 | tee -a "$LOG"
  PENDIENTES=$(tail -80 "$LOG" | grep -oE '[0-9]+ pendientes de procesar' | tail -1 | grep -oE '^[0-9]+')
  if [ "${PENDIENTES:-1}" = "0" ]; then
    echo ">>> [run-extract] sin pendientes -- corrida completa." | tee -a "$LOG"
    node scripts/validar-claves.js 2>&1 | tee -a "$LOG"
    exit 0
  fi
done

echo ">>> [run-extract] se agotaron $MAX_INTENTOS intentos sin terminar -- revisar manualmente (puede ser un expediente realmente atorado, no solo cortes de ventana)." | tee -a "$LOG"
exit 1
