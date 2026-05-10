@echo off
setlocal EnableDelayedExpansion
title Dental Automation - Reconfigure

set PYTHON=

REM ── Per-user install (default install.bat target) ─────────────
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"

REM ── All-users install (Program Files, 64-bit) ────────────────
if "!PYTHON!"=="" if exist "C:\Program Files\Python314\python.exe" set "PYTHON=C:\Program Files\Python314\python.exe"
if "!PYTHON!"=="" if exist "C:\Program Files\Python313\python.exe" set "PYTHON=C:\Program Files\Python313\python.exe"
if "!PYTHON!"=="" if exist "C:\Program Files\Python312\python.exe" set "PYTHON=C:\Program Files\Python312\python.exe"
if "!PYTHON!"=="" if exist "C:\Program Files\Python311\python.exe" set "PYTHON=C:\Program Files\Python311\python.exe"

REM ── All-users install (Program Files (x86), 32-bit) ──────────
if "!PYTHON!"=="" if exist "C:\Program Files (x86)\Python314\python.exe" set "PYTHON=C:\Program Files (x86)\Python314\python.exe"
if "!PYTHON!"=="" if exist "C:\Program Files (x86)\Python313\python.exe" set "PYTHON=C:\Program Files (x86)\Python313\python.exe"
if "!PYTHON!"=="" if exist "C:\Program Files (x86)\Python312\python.exe" set "PYTHON=C:\Program Files (x86)\Python312\python.exe"
if "!PYTHON!"=="" if exist "C:\Program Files (x86)\Python311\python.exe" set "PYTHON=C:\Program Files (x86)\Python311\python.exe"

REM ── Root drive (legacy/custom installs) ──────────────────────
if "!PYTHON!"=="" if exist "C:\Python314\python.exe" set "PYTHON=C:\Python314\python.exe"
if "!PYTHON!"=="" if exist "C:\Python313\python.exe" set "PYTHON=C:\Python313\python.exe"
if "!PYTHON!"=="" if exist "C:\Python312\python.exe" set "PYTHON=C:\Python312\python.exe"
if "!PYTHON!"=="" if exist "C:\Python311\python.exe" set "PYTHON=C:\Python311\python.exe"

REM ── Last resort: PATH ────────────────────────────────────────
if "!PYTHON!"=="" (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON=python"
)
if "!PYTHON!"=="" (
    where py >nul 2>nul
    if not errorlevel 1 set "PYTHON=py"
)

if "!PYTHON!"=="" (
    echo Python not found. Please run install.bat first.
    pause & exit /b 1
)

echo Using Python: !PYTHON!
echo Starting Setup Wizard - please wait...
start "" "!PYTHON!" "%~dp0setup_wizard.py"
echo Waiting for wizard to start...
timeout /t 5 /nobreak >nul
start "" "http://localhost:5001"
echo.
echo If browser shows error, wait 5 more seconds and refresh.
echo Do NOT close this window until setup is complete.
pause
