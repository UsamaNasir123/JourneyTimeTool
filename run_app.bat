@echo off
cd /d "%~dp0"
echo Installing required packages (first time only)...
python -m pip install -r requirements.txt
echo Starting the app. Your browser will open. Close this window to stop the app.
python -m streamlit run app.py
pause
