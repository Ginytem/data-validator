@echo off
title Data Validator Guard
cd /d "%~dp0"
curl -s -o nul -m 3 http://127.0.0.1:8501
if %errorlevel%==0 exit /b 0
python -m streamlit run final.py --server.port 8501 --server.address 0.0.0.0 --browser.gatherUsageStats false --server.headless true
