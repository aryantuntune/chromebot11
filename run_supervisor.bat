@echo off
REM ===================================================================
REM  Supervisor: runs the bot in a loop until EVERY target account has
REM  an OTP (auto-retries failures, cools down on throttling).
REM  Double-click to start. Ctrl+C / close window to stop.
REM  Open run_dashboard.bat in another window to watch live progress.
REM ===================================================================
cd /d "%~dp0"
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if not exist ".venv\Scripts\python.exe" (
  echo Environment missing. Run setup.bat first. & pause & exit /b 1
)

REM --- Settings (edit as needed) ---------------------------------------
REM Row window to process (1-based, inclusive). START=2 & END=0 = whole file.
if not defined BOT_START_ROW set "BOT_START_ROW=2"
if not defined BOT_END_ROW   set "BOT_END_ROW=0"
if not defined BOT_BROWSER       set "BOT_BROWSER=edge"
if not defined BOT_AUTO_CAPTCHA  set "BOT_AUTO_CAPTCHA=1"
if not defined BOT_ADAPTIVE      set "BOT_ADAPTIVE=1"
if not defined BOT_SETTLE_MS     set "BOT_SETTLE_MS=5000"
if not defined BOT_BETWEEN_MS    set "BOT_BETWEEN_MS=5000"
if not defined BOT_MAX_PASSES        set "BOT_MAX_PASSES=6"
if not defined BOT_MAX_STALLS        set "BOT_MAX_STALLS=8"
if not defined BOT_STALL_COOLDOWN_SEC set "BOT_STALL_COOLDOWN_SEC=600"

".venv\Scripts\python.exe" supervisor.py

echo.
echo Supervisor finished. Results in output\OTP_results.xlsx and the input file's OTP column.
pause
