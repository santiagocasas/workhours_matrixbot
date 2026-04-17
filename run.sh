#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [ -f "$PROJECT_DIR/.venv/bin/activate" ]; then
  # Prefer a project-local virtualenv if present.
  . "$PROJECT_DIR/.venv/bin/activate"
fi

export PYTHONPATH="$PROJECT_DIR"
exec python3 -m src
