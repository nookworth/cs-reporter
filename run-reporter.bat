@echo off
REM CS Reporter - Launcher Script for Windows
REM This script activates the virtual environment and runs the reporter

cd /d "%~dp0"
call .venv\Scripts\activate.bat
reporter
