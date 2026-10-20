#!/usr/bin/env bash
# Atualiza o painel a partir do GitHub, direto no servidor (sem PowerShell/Windows).
#
#   sudo /opt/dashboard/deploy-from-git.sh [branch] [--so-preparar]
#
#   branch         padrao: producao
#   --so-preparar  monta a pasta da release e para (nao constroi imagem nem troca o painel)
#
# Precisa de: /opt/dashboard/.env (veja criar-env.sh) e da chave de deploy (somente leitura)
# em /opt/dashboard/.ssh/github_deploy, cadastrada em GitHub > repositorio > Settings > Deploy keys.
# O release.sh faz a troca sem derrubar o painel (blue/green) e volta sozinho se algo falhar.
set -euo pipefail

ROOT="${DASH_ROOT:-/opt/dashboard}"
REPO="${DASH_REPO:-git@github.com:Suporte-Temazec/dashboard_obra.git}"
KEY="${DASH_DEPLOY_KEY:-$ROOT/.ssh/github_deploy}"
BRANCH="producao"
SO_PREPARAR=0
for a in "$@"; do
  case "$a" in
    --so-preparar) SO_PREPARAR=1 ;;
    -*) echo "Opcao desconhecida: $a"; exit 2 ;;
    *) BRANCH="$a" ;;
  esac
done

[ "$(id -u)" = 0 ] || { echo "Rode com sudo."; exit 1; }
[ -f "$ROOT/.env" ] || { echo "Falta $ROOT/.env (rode deploy/criar-env.sh antes)."; exit 1; }
[ -f "$KEY" ] || { echo "Falta a chave de deploy em $KEY."; exit 1; }
command -v git >/dev/null || { echo "git nao instalado (apt-get install -y git)."; exit 1; }

mkdir -p "$ROOT/.ssh" "$ROOT/releases"
chmod 700 "$ROOT/.ssh"
export GIT_SSH_COMMAND="ssh -i $KEY -o IdentitiesOnly=yes -o StrictHostKeyChecking=accept-new -o UserKnownHostsFile=$ROOT/.ssh/known_hosts"

TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

echo "1/3 Baixando $BRANCH do GitHub..."
git clone --quiet --depth 1 --branch "$BRANCH" "$REPO" "$TMP/src"
SHA="$(git -C "$TMP/src" rev-parse --short HEAD)"
ID="$(date +%Y%m%d_%H%M%S)-$SHA"
DEST="$ROOT/releases/$ID"

echo "2/3 Montando a release $ID..."
mkdir -p "$DEST"
# so o que roda no servidor (mesmo conjunto do deploy.ps1)
cp "$TMP/src"/*.py "$DEST/"
for n in requirements.txt Dockerfile .dockerignore; do cp "$TMP/src/$n" "$DEST/"; done
cp -r "$TMP/src/templates" "$TMP/src/static" "$DEST/"
# scripts do servidor
for n in release.sh rollback.sh run-in-active.sh backup.sh deploy-from-git.sh criar-env.sh Caddyfile docker-compose.proxy.yml; do
  [ -f "$TMP/src/deploy/$n" ] && cp "$TMP/src/deploy/$n" "$ROOT/$n"
done
chmod +x "$ROOT"/*.sh
find "$DEST" -name '*.py' -exec python3 -I -m py_compile {} + 2>/dev/null || { echo "Erro de sintaxe Python na release."; rm -rf "$DEST"; exit 1; }
find "$DEST" -name __pycache__ -prune -exec rm -rf {} +

if [ "$SO_PREPARAR" = 1 ]; then
  echo "Modo --so-preparar: release pronta em $DEST (nada foi construido nem trocado)."
  exit 0
fi

echo "3/3 Construindo e trocando sem derrubar (release.sh)..."
"$ROOT/release.sh" "$ID"
