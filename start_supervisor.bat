@echo off
cd /d "%~dp0"
tasklist /FI "IMAGENAME eq python.exe" /FO CSV | findstr /i "supervisor" >nul 2>&1
if not errorlevel 1 (
    echo Supervisor already running. Not starting another.
    pause
    exit /b 0
)
echo Starting supervisor...
.venv\Scripts\python.exe supervisor.py
pause
