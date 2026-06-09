@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   LIXIL Report App - Setup
echo ============================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    echo Please install Python 3.10+ and enable Add Python to PATH.
    echo.
    pause
    exit /b 1
)

echo [OK] Python found:
python --version
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 exit /b %errorlevel%
)

echo Installing packages...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if %errorlevel% neq 0 exit /b %errorlevel%

echo.
echo ============================================
echo   Done.
echo   Local desktop: run_local.bat
echo   Internal server: run_server.bat
echo ============================================
echo.
pause
