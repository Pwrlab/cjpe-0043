#!/bin/bash
# Double-click to start Refilling Tool v2 on macOS. First run installs dependencies (1-3 min).
cd "$(dirname "$0")"
PY=$(command -v python3)
if [ -z "$PY" ] || ! "$PY" -c 'import sys; sys.exit(sys.version_info < (3, 9))'; then
  echo "Python 3.9+ not found. Install it from https://www.python.org/downloads/ ,"
  echo "or download the stand-alone version (no Python needed) from the project's Releases page."
  read -r -p "Press Enter to close"; exit 1
fi
if [ ! -x .venv/bin/refilltool ]; then
  echo "First run: installing (1-3 min)..."
  "$PY" -m venv .venv || { read -r -p "venv failed. Press Enter"; exit 1; }
  ./.venv/bin/python -m pip install -q --upgrade pip --timeout 20 \
    || ./.venv/bin/python -m pip install -q --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
  # PyPI first, Tsinghua mirror as fallback (faster in China)
  ./.venv/bin/python -m pip install -q -e . --timeout 20 \
    || ./.venv/bin/python -m pip install -q -e . -i https://pypi.tuna.tsinghua.edu.cn/simple \
    || { echo "Install failed. Check the network."; read -r -p "Press Enter"; exit 1; }
fi
./.venv/bin/refilltool gui
