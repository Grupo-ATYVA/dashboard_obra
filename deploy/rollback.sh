#!/usr/bin/env bash
# Volta pra release anterior (usa o mesmo caminho sem queda do release.sh).
set -euo pipefail
ROOT="${DASH_ROOT:-/opt/dashboard}"
PREV="$(cat "$ROOT/state/previous_release" 2>/dev/null || true)"
[ -n "$PREV" ] || { echo "Não há release anterior registrada."; exit 1; }
[ -d "$ROOT/releases/$PREV" ] || { echo "Código da release anterior ($PREV) não está mais no servidor."; exit 1; }
echo "Voltando para a release $PREV ..."
exec "$ROOT/release.sh" "$PREV"
