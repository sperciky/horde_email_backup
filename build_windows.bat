@echo off
:: ══════════════════════════════════════════════════════════════
:: build_windows.bat — Build a standalone Windows .exe
::
:: Prerequisites (run once):
::   pip install pyinstaller flask
::
:: Output: dist\HordeExplorer\HordeExplorer.exe
::         (copy the entire dist\HordeExplorer\ folder to the target PC)
:: ══════════════════════════════════════════════════════════════

setlocal
cd /d "%~dp0"

echo Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building HordeExplorer.exe ...
echo.

:: --onedir   : creates a folder (faster startup than --onefile)
:: --windowed : hides the console window (use --console to keep it for debugging)
:: --icon     : optional; add your own .ico file path here
:: --add-data : bundles Flask templates and static files inside the exe folder

pyinstaller ^
  --name HordeExplorer ^
  --onedir ^
  --windowed ^
  --add-data "explorer\templates;templates" ^
  --add-data "explorer\static;static" ^
  --paths "explorer" ^
  --hidden-import flask ^
  --hidden-import werkzeug ^
  --hidden-import jinja2 ^
  --hidden-import sqlite3 ^
  explorer\__main__.py

if errorlevel 1 (
    echo.
    echo BUILD FAILED. See errors above.
    pause
    exit /b 1
)

:: Copy run script into dist folder for convenience
copy run_explorer.bat dist\HordeExplorer\run_explorer.bat >nul

echo.
echo ══════════════════════════════════════════════════
echo  Build successful!
echo  Output: dist\HordeExplorer\HordeExplorer.exe
echo.
echo  To distribute:
echo  1. Copy the entire dist\HordeExplorer\ folder
echo  2. Place your "data\" folder next to HordeExplorer.exe
echo  3. Double-click HordeExplorer.exe to start
echo ══════════════════════════════════════════════════
echo.
pause
