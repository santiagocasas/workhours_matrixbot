#!/usr/bin/env bash
set -euo pipefail

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  cat <<'EOF'
Usage: bash scripts/pull_workbook_from_vps.sh [remote_host] [remote_file] [local_file] [windows_file]

Defaults:
  remote_host   ionosvps
  remote_file   /home/santi/Personal/matrix-instances/work-hours-bot/data/Arbeitszeitkarte 2026.xlsx
  local_file    $HOME/RDM-Software/arbeitszeit/Arbeitszeitkarte 2026.xlsx
  windows_file  /mnt/c/Users/casa_sa/Documents/Admin/Arbeitszeit/2026/Arbeitszeitkarte 2026.xlsx
EOF
  exit 0
fi

REMOTE_HOST="${1:-ionosvps}"
REMOTE_FILE="${2:-/home/santi/Personal/matrix-instances/work-hours-bot/data/Arbeitszeitkarte 2026.xlsx}"
REMOTE_BASENAME="$(basename "$REMOTE_FILE")"
LOCAL_FILE="${3:-$HOME/RDM-Software/arbeitszeit/$REMOTE_BASENAME}"
WINDOWS_FILE="${4:-/mnt/c/Users/casa_sa/Documents/Admin/Arbeitszeit/2026/$REMOTE_BASENAME}"

mkdir -p "$(dirname "$LOCAL_FILE")"
rsync -av "${REMOTE_HOST}:${REMOTE_FILE}" "$LOCAL_FILE"

if [[ -n "$WINDOWS_FILE" ]]; then
  mkdir -p "$(dirname "$WINDOWS_FILE")"
  cp "$LOCAL_FILE" "$WINDOWS_FILE"
fi

printf 'Pulled workbook to %s\n' "$LOCAL_FILE"
if [[ -n "$WINDOWS_FILE" ]]; then
  printf 'Copied workbook to %s\n' "$WINDOWS_FILE"
fi
