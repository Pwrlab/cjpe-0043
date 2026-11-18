@echo off
REM Double-click to start Refilling Tool v2 on Windows. First run installs dependencies (1-3 min).
chcp 65001 >nul
cd /d "%~dp0"

py -3 -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>&1
if %errorlevel%==0 (set "PY=py -3" & goto havepy)
python -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>&1
if %errorlevel%==0 (set "PY=python" & goto havepy)
echo Python 3.9+ not found. Install it from https://www.python.org/downloads/
echo and tick "Add python.exe to PATH", or download the stand-alone version
echo (no Python needed) from the project Releases page.
pause
exit /b 1

:havepy
if exist .venv\Scripts\refilltool.exe goto run
echo First run: installing, 1-3 min...
%PY% -m venv .venv
if errorlevel 1 (echo venv failed. & pause & exit /b 1)
REM PyPI first, Tsinghua mirror as fallback (faster in China)
.venv\Scripts\python -m pip install -q --upgrade pip --timeout 20 || .venv\Scripts\python -m pip install -q --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
.venv\Scripts\python -m pip install -q -e . --timeout 20
if errorlevel 1 .venv\Scripts\python -m pip install -q -e . -i https://pypi.tuna.tsinghua.edu.cn/simple
if not exist .venv\Scripts\refilltool.exe (echo Install failed. Check the network. & pause & exit /b 1)

:run
.venv\Scripts\refilltool gui
pause
