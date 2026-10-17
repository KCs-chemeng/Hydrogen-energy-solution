#!/usr/bin/env bash
set -euo pipefail
cd -- "$(dirname -- "${BASH_SOURCE[0]}")"
"${PSL_PYTHON:-python3}" -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m pip check
.venv/bin/python -c 'import tkinter; print("Tk", tkinter.TkVersion)'
PYTHONDONTWRITEBYTECODE=1 LOCALAPPDATA="${LOCALAPPDATA:-${XDG_DATA_HOME:-$HOME/.local/share}}" .venv/bin/python -m unittest test_core test_v2 -v
printf '%s\n' '설치 완료. 그래픽 데스크톱에서 ./start.sh 실행'
