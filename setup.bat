@echo off
REM CS Reporter - Setup Script for Windows
REM This script sets up the cs-reporter tool for first-time use

echo =========================================
echo   CS Reporter - Setup
echo =========================================
echo.

REM Check if Python 3 is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python 3 is not installed.
    echo Please install Python 3 from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation!
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

REM Create virtual environment
echo Creating virtual environment...
if exist ".venv" (
    echo   Virtual environment already exists, skipping...
) else (
    python -m venv .venv
    echo   Virtual environment created
)
echo.

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat
echo   Activated
echo.

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip --quiet
echo   pip upgraded
echo.

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt --quiet
echo   Dependencies installed
echo.

REM Install package in editable mode
echo Installing cs-reporter...
pip install -e . --quiet
echo   cs-reporter installed
echo.

REM Create launcher script
echo Creating launcher script...
(
echo @echo off
echo REM CS Reporter - Launcher Script for Windows
echo REM This script activates the virtual environment and runs the reporter
echo.
echo cd /d "%%~dp0"
echo call .venv\Scripts\activate.bat
echo reporter
) > run-reporter.bat

echo   Launcher script created
echo.

echo =========================================
echo   Setup Complete!
echo =========================================
echo.
echo To run the reporter, use one of these commands:
echo.
echo   Option 1 (Recommended):
echo     run-reporter.bat
echo.
echo   Option 2:
echo     .venv\Scripts\activate.bat
echo     reporter
echo.
echo The reporter will prompt you to select your Excel files.
echo Generated reports will be saved in the 'output' folder.
echo.
pause
