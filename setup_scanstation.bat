@echo off
setlocal enabledelayedexpansion
title Scan Station — Setup
color 0A

:: Keep window open on any error
if "%1"=="ELEVATED" goto :main
echo Relaunching as administrator...
powershell -Command "Start-Process cmd -ArgumentList '/k \"%~f0\" ELEVATED' -Verb RunAs"
exit /b

:main
echo.
echo  ============================================
echo   LEGO Scan Station — Automatic Setup
echo  ============================================
echo.

:: ── Check / Install Python ────────────────────────────────────────────────
echo  [1/5] Checking Python...
py --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Python not found. Downloading Python 3.13...
    curl -L -o "%TEMP%\python_installer.exe" "https://www.python.org/ftp/python/3.13.0/python-3.13.0-amd64.exe"
    if %errorlevel% neq 0 (
        echo  [!] Download failed. Please install Python manually from https://python.org
        goto :end
    )
    echo  [*] Installing Python 3.13...
    "%TEMP%\python_installer.exe" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1 Include_launcher=1
    echo  [+] Python installed.
) else (
    for /f "tokens=*" %%v in ('py --version 2^>^&1') do echo  [+] Found %%v
)

:: ── Upgrade pip ───────────────────────────────────────────────────────────
echo.
echo  [2/5] Upgrading pip...
py -m pip install --upgrade pip --quiet
echo  [+] pip up to date.

:: ── Install packages ──────────────────────────────────────────────────────
echo.
echo  [3/5] Installing Python packages...
echo.

set PACKAGES=PyQt5 requests requests-oauthlib Pillow numpy scipy opencv-python python-dotenv

for %%p in (%PACKAGES%) do (
    echo  [*] %%p...
    py -m pip install %%p --quiet
    if !errorlevel! neq 0 (
        echo  [!] %%p failed — retrying...
        py -m pip install %%p
    ) else (
        echo  [+] %%p OK
    )
)

:: ── DroidCam reminder ─────────────────────────────────────────────────────
echo.
echo  [4/5] Optional: DroidCam Pro PC client
echo    https://www.dev47apps.com/
echo.

:: ── Files and folders ─────────────────────────────────────────────────────
echo  [5/5] Checking files and folders...
echo.

if not exist ".env" (
    (
        echo CONSUMER_KEY=your_bricklink_consumer_key
        echo CONSUMER_SECRET=your_bricklink_consumer_secret
        echo ACCESS_TOKEN=your_bricklink_access_token
        echo TOKEN_SECRET=your_bricklink_token_secret
    ) > .env
    echo  [+] .env created - fill in BrickLink API credentials
) else ( echo  [+] .env found - keeping existing )

if not exist "station.cfg" (
    (
        echo [camera]
        echo index = 1
        echo backend = DSHOW
        echo [grid]
        echo enabled = false
        echo cols = 8
        echo rows = 6
    ) > station.cfg
    echo  [+] station.cfg created
) else ( echo  [+] station.cfg found - keeping existing )

if exist "scan-gui.py"   ( echo  [+] scan-gui.py found
) else ( echo  [!] WARNING: scan-gui.py missing )

if exist "scan-heads.py" ( echo  [+] scan-heads.py found
) else ( echo  [!] WARNING: scan-heads.py missing )

if exist "parts.csv"           ( echo  [+] parts.csv found
) else ( echo  [~] parts.csv missing - optional )

if exist "inventory_parts.csv" ( echo  [+] inventory_parts.csv found
) else ( echo  [~] inventory_parts.csv missing - optional )

if not exist "scans"       mkdir scans
if not exist "reports"     mkdir reports
if not exist "image_cache" mkdir image_cache
echo  [+] Folders ready

:: ── Verify ────────────────────────────────────────────────────────────────
echo.
echo  Verifying packages...
py -c "import PyQt5; import cv2; import numpy; import scipy; import requests; import PIL; print('  [+] All packages OK')"
if %errorlevel% neq 0 (
    echo  [!] Some packages failed - see errors above
)

:: ── Done ──────────────────────────────────────────────────────────────────
echo.
echo  ============================================
echo   Done! Run with:  py scan-gui.py
echo  ============================================
echo.

:end
echo.
echo  Press any key to close...
pause >nul
endlocal
