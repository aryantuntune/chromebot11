@echo off
REM ===================================================================
REM  Live dashboard for the HP Gas OTP bot (read-only viewer).
REM  Double-click while the bot/supervisor is running to watch progress.
REM  Press Ctrl+C to close the dashboard (does NOT stop the bot).
REM ===================================================================
cd /d "%~dp0"
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if not exist ".venv\Scripts\python.exe" (
  echo Environment missing. Run setup.bat first. & pause & exit /b 1
)

REM Match the row window the bot is processing (leave blank = whole file).
if not defined BOT_START_ROW set "BOT_START_ROW=247"
if not defined BOT_END_ROW   set "BOT_END_ROW=312"

".venv\Scripts\python.exe" dashboard.py
