@echo off
setlocal
title TSG Automate - Install Requirements (No venv)

REM Go to the folder where this .bat lives
set "ROOT=%~dp0"
pushd "%ROOT%"

REM Create requirements.txt if missing
if not exist "requirements.txt" (
  echo Creating requirements.txt...
  > requirements.txt echo PySide6
  >>requirements.txt echo selenium>=4.10
  >>requirements.txt echo webdriver-manager
  >>requirements.txt echo pdfplumber
  >>requirements.txt echo pandas
  >>requirements.txt echo openpyxl
  >>requirements.txt echo pywin32
)

REM Find Python (prefer 'py', else 'python')
where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
  set "PY_CMD=py -3"
) else (
  where python >nul 2>&1
  if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python not found. Install Python 3.10+ and make sure it's on PATH.
    echo Download: https://www.python.org/downloads/windows/
    echo.
    pause
    exit /b 1
  )
  set "PY_CMD=python"
)

echo.
echo === Upgrading pip (user) ===
%PY_CMD% -m pip install --upgrade --user pip
if %ERRORLEVEL% NEQ 0 (
  echo [WARN] Could not upgrade pip. Continuing...
)

echo.
echo === Installing requirements (user site) ===
%PY_CMD% -m pip install --user -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
  echo [ERROR] Failed to install one or more packages.
  echo If you're behind a proxy, configure pip proxy settings and try again.
  echo.
  pause
  exit /b 1
)

echo.
echo ✅ Packages installed (user site).
echo.
echo To run the app:
echo   %PY_CMD% -u TSG_automate_app.py
echo.
pause
exit /b 0
