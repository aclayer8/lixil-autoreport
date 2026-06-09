@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   LIXIL Report App - Local Desktop
echo ============================================
echo.
echo URL: http://127.0.0.1:5000
echo Press Ctrl+C to stop.
echo.

set LIXIL_APP_HOST=127.0.0.1
set LIXIL_APP_PORT=5000
set LIXIL_AUTO_OPEN_BROWSER=1
set LIXIL_SHOW_ERROR_TRACE=1

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py
) else (
    python app.py
)
pause
