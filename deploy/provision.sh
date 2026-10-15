#!/usr/bin/env bash
# Prepara um servidor Ubuntu 22.04/24.04 NOVO (rodar UMA vez, como root):
#   sudo bash provision.sh
set -euo pipefail
[ "$(id -u)" = 0 ] || { echo "Rode como root (sudo)."; exit 1; }
ROOT=/opt/dashboard

apt-get update
DEBIAN_FRONTEND=noninteractive apt-get install -y ca-certificates curl ufw fail2ban unattended-upgrades rsync cron

# Docker oficial
if ! command -v docker >/dev/null 2>&1; then
  curl -fsSL https://get.docker.com | sh
fi
systemctl enable --now docker

timedatectl set-timezone America/Sao_Paulo

# pastas (dados/ pertence ao usuário 1000 do container)
mkdir -p "$ROOT"/{dados,logs,releases,backups,state/caddy}
chown -R 1000:1000 "$ROOT/dados" "$ROOT/logs"
chmod 700 "$ROOT" "$ROOT/dados" "$ROOT/backups"

# firewall: só SSH, HTTP e HTTPS
ufw allow OpenSSH
ufw allow 80/tcp
ufw allow 443/tcp
ufw --force enable

# pouca memória? cria 2 GB de swap (a atualização sobe duas cópias do painel por instantes)
if [ "$(awk '/MemTotal/ {print int($2/1024)}' /proc/meminfo)" -lt 5000 ] && [ ! -f /swapfile ]; then
  fallocate -l 2G /swapfile && chmod 600 /swapfile && mkswap /swapfile && swapon /swapfile
  echo '/swapfile none swap sw 0 0' >> /etc/fstab
fi

systemctl enable --now fail2ban unattended-upgrades

echo
echo "Servidor pronto. Próximos passos estão em deploy/LEIA_ME_LINUX.md (seção 3)."
echo "Pasta do painel: $ROOT"
