@echo off
REM ===================================================================
REM  Start the HP Gas sign-in bot (Windows).
REM  Double-click this file after running setup.bat once.
REM  To STOP early: close this window, or press Ctrl+C.
REM  Progress is saved to the "output" folder after every row.
REM ===================================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
  echo.
  echo The environment is missing. Please run setup.bat first.
  pause
  exit /b 1
)

call .venv\Scripts\activate.bat
python run_bot.py

echo.
echo Bot finished (or was stopped). Output is in the "output" folder.
pause
