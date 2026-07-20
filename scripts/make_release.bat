@echo off
REM ============================================================
REM  Build single-file exe and publish it to release folder
REM  Usage:  scripts\make_release.bat [version]
REM  Example: scripts\make_release.bat v1.0.0
REM  Output: release\cmb_pdf2excel-<version>-win-x64.exe
REM ============================================================
setlocal
cd /d "%~dp0\.."

set VERSION=%~1
if "%VERSION%"=="" set VERSION=v1.1.0

echo === [1/3] Building worker.exe ===
call scripts\build_worker.bat || exit /b 1

echo.
echo === [2/3] Building cmb_pdf2excel.exe (embed worker.exe) ===
call scripts\build_gui.bat || exit /b 1

echo.
echo === [3/3] Publishing release ===
if not exist "release" mkdir release
set OUT_NAME=cmb_pdf2excel-%VERSION%-win-x64.exe

copy /y "dist\cmb_pdf2excel.exe" "release\%OUT_NAME%" >nul

if exist "release\%OUT_NAME%" (
    echo.
    echo [OK] release\%OUT_NAME% created
    echo Upload this single exe to GitHub Releases.
) else (
    echo [ERROR] publishing failed
    exit /b 1
)
endlocal
