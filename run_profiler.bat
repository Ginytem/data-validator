@echo off
cd /d "%~dp0"
echo Starting Data Profiler... browser will open automatically.
python -m streamlit run app.py --browser.gatherUsageStats false
pause
