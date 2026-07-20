@echo off
REM ============================================================
REM  Build worker.exe (Python console worker, PyInstaller)
REM  Usage:  scripts\build_worker.bat
REM  Prereq: pip install pymupdf fonttools openpyxl pyinstaller
REM ============================================================
setlocal
cd /d "%~dp0\.."

if not exist "src\worker.py" (
    echo [ERROR] src\worker.py not found
    exit /b 1
)

cd src
python -m PyInstaller -y --onefile --console --name worker --clean ^
    --exclude-module pandas --exclude-module numpy ^
    --exclude-module tkinter --exclude-module _tkinter ^
    --exclude-module matplotlib --exclude-module scipy ^
    --exclude-module sqlalchemy --exclude-module PIL --exclude-module Pillow ^
    --exclude-module pythonnet --exclude-module clr --exclude-module clr_loader ^
    --exclude-module pycparser --exclude-module cryptography --exclude-module cffi ^
    --exclude-module webview ^
    --distpath "..\dist" --workpath "..\build" --specpath "..\build" ^
    worker.py
cd ..

if exist "dist\worker.exe" (
    echo.
    echo [OK] dist\worker.exe built
) else (
    echo [ERROR] build failed
    exit /b 1
)
endlocal
