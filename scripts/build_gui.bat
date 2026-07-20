@echo off
REM ============================================================
REM  Build single-file exe (C# WinForms + embedded worker.exe)
REM  Usage:  scripts\build_gui.bat
REM  Prereq: dist\worker.exe must exist (run build_worker.bat first)
REM ============================================================
setlocal
cd /d "%~dp0\.."

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" (
    echo [ERROR] C# compiler not found: %CSC%
    exit /b 1
)
if not exist "src\gui.cs" (
    echo [ERROR] src\gui.cs not found
    exit /b 1
)
if not exist "dist\worker.exe" (
    echo [ERROR] dist\worker.exe not found - run build_worker.bat first
    exit /b 1
)

REM 图标缺失时自动生成
if not exist "assets\app.ico" (
    echo [INFO] Generating assets\app.ico ...
    if not exist "assets" mkdir assets
    "%CSC%" /nologo /target:exe /out:"scripts\make_icon.exe" /r:System.Drawing.dll scripts\make_icon.cs
    "scripts\make_icon.exe" "assets\app.ico"
)

if not exist "dist" mkdir dist
"%CSC%" /target:winexe /out:"dist\cmb_pdf2excel.exe" /win32icon:"assets\app.ico" /resource:"dist\worker.exe" src\gui.cs

if exist "dist\cmb_pdf2excel.exe" (
    echo.
    echo [OK] cmb_pdf2excel.exe built, worker.exe embedded
) else (
    echo [ERROR] build failed
    exit /b 1
)
endlocal
