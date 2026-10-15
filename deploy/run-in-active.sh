#!/usr/bin/env bash
# Roda um comando dentro do container que está no ar. Ex.:
#   ./run-in-active.sh python carta_convite.py
ROOT="${DASH_ROOT:-/opt/dashboard}"
COR="$(cat "$ROOT/state/active_color" 2>/dev/null || echo blue)"
exec docker exec "dash-$COR" "$@"
