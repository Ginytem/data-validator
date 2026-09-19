@echo off
title Data Validator Guard
cd /d "%~dp0"
set "PY=%APPDATA%\uv\python\cpython-3.12.14-windows-x86_64-none\python.exe"
if not exist "%PY%" set "PY=python"
curl -s -o nul -m 3 http://127.0.0.1:8501
if %errorlevel%==0 exit /b 0
"%PY%" -m streamlit run final.py --server.port 8501 --server.address 0.0.0.0 --browser.gatherUsageStats false --server.headless true --server.maxUploadSize 20
