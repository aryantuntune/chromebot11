@echo off
REM ===================================================================
REM  One-time setup for the HP Gas OTP bot (Windows).
REM  Double-click this file. It finds Python, builds a private
REM  environment, and installs everything (incl. the CAPTCHA solver).
REM  No coding or AI tools required -- just run it.
REM ===================================================================
setlocal
cd /d "%~dp0"

echo.
echo [1/4] Locating Python 3.9+ ...
set "PYEXE="
REM Prefer the 'py' launcher (works even when 'python' is shadowed by the
REM Microsoft Store stub); fall back to 'python'.
py -3 --version >nul 2>&1 && set "PYEXE=py -3"
if not defined PYEXE (
  python --version >nul 2>&1 && set "PYEXE=python"
)
if not defined PYEXE (
  echo.
  echo ERROR: Python 3.9+ was not found.
  echo Install it from https://python.org/downloads and tick
  echo "Add Python to PATH" during install, then run this again.
  pause
  exit /b 1
)
echo       using: %PYEXE%

echo [2/4] Creating the virtual environment (.venv) ...
%PYEXE% -m venv .venv
if errorlevel 1 (
  echo ERROR: could not create the environment. & pause & exit /b 1
)

echo [3/4] Installing dependencies (playwright, openpyxl, ddddocr CAPTCHA OCR) ...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo ERROR: dependency install failed. Check your internet connection. & pause & exit /b 1
)

echo [4/4] Registering Google Chrome for automation ...
".venv\Scripts\python.exe" -m playwright install chrome

echo.
echo ===================================================================
echo  Setup complete!
echo  1. Put your TWO .xlsx files in the "input" folder:
echo       - the master credentials file (many columns)
echo       - the targets file (one column of consumer IDs)
echo  2. Double-click run.bat to start the bot.
echo  See HANDOFF.md for full details.
echo ===================================================================
pause
endlocal
