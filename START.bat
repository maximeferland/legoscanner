@echo off
cd /d "%~dp0"
py scan-gui.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: scan-gui.py crashed with code %ERRORLEVEL%
    pause
)