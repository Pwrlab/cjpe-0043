@echo off
REM Double-click to start Refilling Tool v2 on Windows (first run installs dependencies).
cd /d "%~dp0"
if not exist .venv (
  echo First run: creating virtual environment and installing dependencies...
  python -m venv .venv || (echo Python 3 not found. Install it from https://www.python.org & pause & exit /b 1)
  .venv\Scripts\python -m pip install --upgrade pip -q
  .venv\Scripts\python -m pip install -e . -q || (echo Install failed. & pause & exit /b 1)
)
.venv\Scripts\refilltool gui
pause
