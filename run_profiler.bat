@echo off
cd /d "%~dp0"
set "PY=%APPDATA%\uv\python\cpython-3.12.14-windows-x86_64-none\python.exe"
if not exist "%PY%" set "PY=python"
echo Using Python: %PY%
"%PY%" -m streamlit run app.py --browser.gatherUsageStats false
pause
