@echo off
REM ===================================================================
REM  One-time setup for the HP Gas sign-in bot (Windows).
REM  Double-click this file, or run it from a Command Prompt.
REM  It creates a private Python environment and installs everything.
REM ===================================================================
cd /d "%~dp0"

echo.
echo [1/4] Creating the Python virtual environment (.venv)...
python -m venv .venv
if errorlevel 1 (
  echo.
  echo ERROR: Could not create the environment.
  echo Make sure Python 3.9+ is installed from https://python.org
  echo and that you ticked "Add Python to PATH" during install.
  pause
  exit /b 1
)

echo [2/4] Activating the environment...
call .venv\Scripts\activate.bat

echo [3/4] Installing dependencies...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo ERROR: Dependency installation failed. Check your internet connection.
  pause
  exit /b 1
)

echo [4/4] Registering Google Chrome for automation...
python -m playwright install chrome

echo.
echo ===================================================================
echo  Setup complete!
echo  1. Put your .xlsx file (columns: conno, id, password, code) in the
echo     "input" folder.
echo  2. Double-click run.bat to start the bot.
echo ===================================================================
pause
