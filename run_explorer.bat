@echo off
:: ══════════════════════════════════════════════════════════════
:: run_explorer.bat — Windows launcher for the offline explorer
:: Double-click to start the local web server and open a browser.
:: ══════════════════════════════════════════════════════════════

setlocal
cd /d "%~dp0"

:: Check for Python
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.9+ and add it to PATH.
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Install Flask if needed
echo Checking dependencies...
pip install -q flask

:: Check that backup data exists
if not exist "data\index.db" (
    echo.
    echo ERROR: No backup data found at data\index.db
    echo Please run run_backup.bat first to create your email archive.
    echo.
    pause
    exit /b 1
)

echo.
echo Starting Horde Email Explorer at http://localhost:5000
echo Your browser will open automatically.
echo Press Ctrl+C to stop the server.
echo.

python explorer\app.py --data data --port 5000

pause
