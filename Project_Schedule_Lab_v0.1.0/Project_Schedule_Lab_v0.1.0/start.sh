#!/usr/bin/env bash
set -euo pipefail
cd -- "$(dirname -- "${BASH_SOURCE[0]}")"
if [ ! -x .venv/bin/python ]; then
  printf '%s\n' '먼저 ./install.sh를 실행하세요.' >&2
  exit 1
fi
export LOCALAPPDATA="${LOCALAPPDATA:-${XDG_DATA_HOME:-$HOME/.local/share}}"
export PYTHONDONTWRITEBYTECODE=1
exec .venv/bin/python app.py
