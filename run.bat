@echo off
REM ===================================================================
REM  Start the HP Gas OTP bot (Windows).
REM  Double-click this file after running setup.bat once.
REM  The bot solves the CAPTCHA itself, paces around the portal's
REM  throttle, and auto-retries failures -- it runs hands-free.
REM  To STOP early: close this window, or press Ctrl+C.
REM  Progress is saved to the "output" folder after every account, and
REM  re-running resumes (skips done, retries only what failed).
REM ===================================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo.
  echo The environment is missing. Please run setup.bat first.
  pause
  exit /b 1
)

REM Sensible defaults (override by setting them before launching if you like).
if not defined BOT_BROWSER set "BOT_BROWSER=edge"
if not defined BOT_SETTLE_MS set "BOT_SETTLE_MS=5000"
if not defined BOT_BETWEEN_MS set "BOT_BETWEEN_MS=5000"

".venv\Scripts\python.exe" run_bot.py

echo.
echo Bot finished (or was stopped). Results are in the "output" folder:
echo   - Captured_OTPs.xlsx  (clean: Consumer ID ^| otp ^| value)
echo   - OTP_results.xlsx    (full log incl. any retryable failures)
pause
