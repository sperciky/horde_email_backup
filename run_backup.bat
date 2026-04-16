@echo off
:: ══════════════════════════════════════════════════════════════
:: run_backup.bat — Windows launcher for the email backup tool
:: Double-click this file to start a backup, OR open a command
:: prompt in this directory and run it with optional arguments.
::
:: Usage:
::   run_backup.bat                   <- incremental backup
::   run_backup.bat --full            <- full re-download
::   run_backup.bat --folder INBOX    <- backup one folder only
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

:: Install/check dependencies (only Flask is needed for explorer; backup uses stdlib)
echo Checking dependencies...
pip install -q -r requirements.txt

echo.
echo Starting email backup...
echo (Press Ctrl+C to stop at any time — progress is saved)
echo.

python backup.py %*

echo.
if errorlevel 1 (
    echo Backup finished with errors. Check backup.log for details.
) else (
    echo Backup complete!
)

pause
