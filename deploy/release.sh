#!/usr/bin/env bash
# Atualiza o painel SEM derrubar (blue/green). Roda no servidor.
#
#   ./release.sh <id-da-release>      (a pasta releases/<id> já deve ter o código)
#
# Fluxo: monta a imagem -> sobe o container NOVO ao lado do atual -> testa de verdade
# (carrega os dados do painel) -> troca o proxy -> aposenta o antigo.
# Qualquer falha ANTES da troca deixa o container antigo atendendo: nada cai.
# Se a troca for feita e o endereço público não responder, volta sozinho pro antigo.
set -euo pipefail

ROOT="${DASH_ROOT:-/opt/dashboard}"
ID="${1:?uso: release.sh <id-da-release>}"
SRC="$ROOT/releases/$ID"
STATE="$ROOT/state"
MEM="${DASH_MEM:-2500m}"
DOMAIN="$(grep -E '^DASH_DOMAIN=' "$ROOT/.env" 2>/dev/null | head -1 | cut -d= -f2- | tr -d '\r' || true)"

log() { echo "[$(date +%H:%M:%S)] $*"; }

mkdir -p "$STATE/caddy" "$ROOT/logs"
exec 9>"$STATE/.lock"
flock -n 9 || { echo "Já tem uma atualização em andamento."; exit 1; }
[ -d "$SRC" ] || { echo "Release não encontrada: $SRC"; exit 1; }
[ -f "$ROOT/.env" ] || { echo "Falta o arquivo $ROOT/.env"; exit 1; }

OLD="$(cat "$STATE/active_color" 2>/dev/null || echo none)"
if [ "$OLD" = blue ]; then NEW=green; else NEW=blue; fi
IMG="dash:$ID"

falhou() {
  log "FALHOU: $1 — a versão antiga continua no ar."
  docker logs --tail 80 "dash-$NEW" 2>&1 || true
  docker rm -f "dash-$NEW" >/dev/null 2>&1 || true
  exit 1
}

upstream_para() {  # $1 = blue|green
  printf 'reverse_proxy dash-%s:5070 {\n\tlb_try_duration 20s\n\tlb_try_interval 250ms\n}\n' "$1" > "$STATE/caddy/upstream.caddy.tmp"
  mv -f "$STATE/caddy/upstream.caddy.tmp" "$STATE/caddy/upstream.caddy"
  if docker ps -q -f name='^caddy$' | grep -q .; then
    docker exec caddy caddy reload --config /etc/caddy/Caddyfile --force >/dev/null
  else
    log "AVISO: container do proxy (caddy) não está rodando."
  fi
}

log "1/6 Montando a imagem $IMG (camadas em cache: só o código novo é refeito)..."
docker build --build-arg BUILD_ID="$ID" -t "$IMG" "$SRC" || { log "FALHOU na montagem da imagem — nada mudou no ar."; exit 1; }

log "2/6 Subindo a versão nova ($NEW) ao lado da atual ($OLD)..."
docker rm -f "dash-$NEW" >/dev/null 2>&1 || true
docker run -d --name "dash-$NEW" --network dashnet --restart unless-stopped \
  --log-opt max-size=50m --log-opt max-file=5 \
  --env-file "$ROOT/.env" -e DASH_BUILD_ID="$ID" \
  -v "$ROOT/dados:/app/dados" -v "$ROOT/logs:/app/logs" \
  --shm-size=1g --memory="$MEM" "$IMG" >/dev/null

log "3/6 Esperando a versão nova responder..."
ok=0
for _ in $(seq 1 60); do
  if docker exec "dash-$NEW" curl -fsS "http://127.0.0.1:5070/api/status" 2>/dev/null | grep -q "$ID"; then ok=1; break; fi
  sleep 2
done
[ "$ok" = 1 ] || falhou "não respondeu em 2 minutos"

log "4/6 Teste de verdade: a versão nova precisa carregar os dados do painel..."
docker exec -i "dash-$NEW" python - "$ID" <<'PY' || falhou "a versão nova não conseguiu carregar os dados"
import json, sys, urllib.request
d = json.load(urllib.request.urlopen("http://127.0.0.1:5070/api/status?deep=1", timeout=900))
ok = d.get("ok") and d.get("build") == sys.argv[1] and (d.get("deep") or {}).get("ok")
print("   status:", {k: d.get(k) for k in ("ok", "build")}, "| deep:", d.get("deep"))
sys.exit(0 if ok else 1)
PY
docker exec "dash-$NEW" curl -fsS -o /dev/null "http://127.0.0.1:5070/login" || falhou "tela de login não abriu"

log "5/6 Trocando o proxy para $NEW..."
upstream_para "$NEW"
if [ -n "$DOMAIN" ]; then
  pub=0
  for _ in $(seq 1 15); do
    if curl -fsSk --resolve "$DOMAIN:443:127.0.0.1" "https://$DOMAIN/api/status" 2>/dev/null | grep -q "$ID"; then pub=1; break; fi
    sleep 2
  done
  if [ "$pub" != 1 ]; then
    log "Endereço público não respondeu com a versão nova — voltando pra $OLD."
    [ "$OLD" != none ] && upstream_para "$OLD"
    falhou "o endereço público não respondeu após a troca"
  fi
fi
echo "$NEW" > "$STATE/active_color"
[ -f "$STATE/current_release" ] && cp "$STATE/current_release" "$STATE/previous_release" || true
echo "$ID" > "$STATE/current_release"

if [ "$OLD" != none ]; then
  log "6/6 Versão nova no ar. Aposentando a antiga ($OLD) depois de uns segundos pra terminar o que estiver em andamento..."
  sleep 20
  docker stop -t 120 "dash-$OLD" >/dev/null 2>&1 || true
  docker rm "dash-$OLD" >/dev/null 2>&1 || true
else
  log "6/6 Primeira versão no ar."
fi

# mantém só as 5 últimas releases e 4 últimas imagens (rollback rápido continua possível)
ls -1dt "$ROOT"/releases/*/ 2>/dev/null | tail -n +6 | xargs -r rm -rf
docker images dash --format '{{.Tag}}' | sort -r | tail -n +5 | xargs -r -I{} docker rmi "dash:{}" >/dev/null 2>&1 || true

log "OK — release $ID no ar (cor $NEW)."
