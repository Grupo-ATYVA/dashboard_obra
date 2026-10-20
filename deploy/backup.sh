#!/usr/bin/env bash
# Backup diário da pasta dados/ (dados financeiros + usuários + certificado).
# Guarda 14 dias em /opt/dashboard/backups. Pra cópia FORA do servidor, descomente o rclone no fim.
set -euo pipefail
ROOT="${DASH_ROOT:-/opt/dashboard}"
DEST="$ROOT/backups"
mkdir -p "$DEST"
STAMP="$(date +%Y%m%d_%H%M)"
umask 077
tar -czf "$DEST/dados_$STAMP.tgz" -C "$ROOT" \
  --exclude='dados/synergia/.melius_profile*' \
  --exclude='dados/synergia/.sienge_profile*' \
  --exclude='dados/synergia/dashboard_data_cache.*' \
  --exclude='*.nwd' \
  dados
find "$DEST" -name 'dados_*.tgz' -mtime +14 -delete
echo "Backup ok: $DEST/dados_$STAMP.tgz ($(du -h "$DEST/dados_$STAMP.tgz" | cut -f1))"
# rclone copy "$DEST/dados_$STAMP.tgz" remoto:backup-dashboard/   # (opcional) cópia fora do servidor
