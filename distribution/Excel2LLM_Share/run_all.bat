@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul && set "PS_EXE=pwsh"

if "%~1"=="" goto :usage
if /I "%~1"=="-h" goto :usage
if /I "%~1"=="--help" goto :usage
if /I "%~1"=="/?" goto :usage

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\run_all.ps1" %*
exit /b %errorlevel%

:usage
echo Usage: run_all.bat "C:\path\to\book.xlsx" [-Verify] [extract options]
echo.
echo Common options:
echo   -Verify
echo   -OutputDir "C:\path\to\output"
echo   -RedactPaths
echo   -Sheets Summary,Calc
echo.
echo See: docs\guides\MANUAL.md or docs\guides\USER_GUIDE.md
exit /b 1
