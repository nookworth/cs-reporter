@echo off
REM CS Reporter Web UI Launcher

echo Starting CS Reporter Web Interface...
echo.
echo The web interface will open in your browser.
echo Press Ctrl+C to stop the server.
echo.

python -m streamlit run app_v3.py

pause
