@echo off
cd /d "%~dp0"
echo Starting Data Validator... browser will open automatically.
python -m streamlit run final.py --browser.gatherUsageStats false
pause
