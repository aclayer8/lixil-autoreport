@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   LIXIL Report App - Internal Server
echo ============================================
echo.
echo Server URL: http://%COMPUTERNAME%:5000
echo Bind: 0.0.0.0:5000
echo Press Ctrl+C to stop.
echo.

set LIXIL_APP_HOST=0.0.0.0
set LIXIL_APP_PORT=5000
set LIXIL_AUTO_OPEN_BROWSER=0
set LIXIL_SHOW_ERROR_TRACE=0

if exist ".venv\Scripts\waitress-serve.exe" (
    ".venv\Scripts\waitress-serve.exe" --host=%LIXIL_APP_HOST% --port=%LIXIL_APP_PORT% app:app
) else if exist ".venv\Scripts\python.exe" (
    echo [WARN] waitress is not installed. Run setup.bat for the recommended server runner.
    ".venv\Scripts\python.exe" app.py
) else (
    echo [WARN] .venv was not found. Run setup.bat before customer deployment.
    python app.py
)
pause
