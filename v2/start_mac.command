#!/bin/bash
# Double-click to start Refilling Tool v2 on macOS (first run installs dependencies).
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "First run: creating virtual environment and installing dependencies..."
  python3 -m venv .venv || { echo "Python 3 not found. Install it from https://www.python.org"; read -r; exit 1; }
  ./.venv/bin/python -m pip install --upgrade pip -q
  ./.venv/bin/python -m pip install -e . -q || { echo "Install failed."; read -r; exit 1; }
fi
./.venv/bin/refilltool gui
