@echo off
cd /d "%~dp0"
echo Installing / checking libraries (first run only)...
python -m pip install -q -r requirements.txt
if errorlevel 1 (echo Python or pip not found. Install Python from python.org and tick "Add to PATH". & pause & exit /b 1)
python -m streamlit run app.py --server.address 0.0.0.0 --browser.gatherUsageStats false
pause
