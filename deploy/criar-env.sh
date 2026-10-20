#!/usr/bin/env bash
# Cria /opt/dashboard/.env no servidor com chaves aleatorias, SEM digitar segredos no chat.
#
#   sudo ./criar-env.sh engenharia.grupoatyva.com.br suportetema@temazec.com
#
# As senhas das integracoes (AutoDoc, Melius, Sienge, certificado...) ficam em branco: preencha
# depois, no proprio servidor, com   sudo nano /opt/dashboard/.env   e salve (Ctrl+O, Enter, Ctrl+X).
set -euo pipefail

ROOT="${DASH_ROOT:-/opt/dashboard}"
DOMINIO="${1:?uso: criar-env.sh <dominio> <email>}"
EMAIL="${2:?uso: criar-env.sh <dominio> <email>}"

[ "$(id -u)" = 0 ] || { echo "Rode com sudo."; exit 1; }
if [ -e "$ROOT/.env" ]; then
  echo "$ROOT/.env ja existe — nao vou sobrescrever. Edite com: sudo nano $ROOT/.env"
  exit 1
fi
case "$DOMINIO" in *[!a-zA-Z0-9.-]*|"") echo "Dominio invalido."; exit 1 ;; esac
case "$EMAIL" in *[!a-zA-Z0-9.@_+-]*|"") echo "E-mail invalido."; exit 1 ;; esac

mkdir -p "$ROOT"
umask 077
SEGREDO="$(openssl rand -hex 32)"
TOKEN="$(openssl rand -hex 16)"

cat > "$ROOT/.env" <<EOF
# Gerado por criar-env.sh. Permissao 600 (so root le). NUNCA versionar nem enviar por chat/e-mail.
DASH_DOMAIN=$DOMINIO
ACME_EMAIL=$EMAIL
DASHBOARD_HTTPS=1
DASHBOARD_FRAME_ANCESTORS=https://temazeccombr.sharepoint.com
DASHBOARD_SESSION_HOURS=12

# assina a sessao dos logins e protege a senha de e-mail salva na Carta Convite
DASHBOARD_SECRET_KEY=$SEGREDO
# so vale no PRIMEIRO acesso (criar o administrador); depois de criar, apague esta linha
DASHBOARD_SETUP_TOKEN=$TOKEN

# integracoes: preencher no servidor (sudo nano $ROOT/.env)
AUTODOC_EMAIL=
AUTODOC_SENHA=
CERT_PFX_SENHA=
MELIUS_USER=
MELIUS_PASSWORD=
MELIUS_SITE_SENHA=
MELIUS_TOKEN=
DIRECTD_API_TOKEN=

GUNICORN_THREADS=16
GUNICORN_TIMEOUT=900
EOF
chmod 600 "$ROOT/.env"
chown root:root "$ROOT/.env"

echo
echo ".env criado em $ROOT/.env (permissao 600)."
echo "Link do PRIMEIRO acesso (cria o administrador; guarde, nao compartilhe):"
echo "  https://$DOMINIO/setup-admin?token=$TOKEN"
