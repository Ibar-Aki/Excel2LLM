@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul && set "PS_EXE=pwsh"

if "%~1"=="" goto :usage
if /I "%~1"=="-h" goto :usage
if /I "%~1"=="--help" goto :usage
if /I "%~1"=="/?" goto :usage

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\pack_for_llm.ps1" %*
exit /b %errorlevel%

:usage
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\show_usage.ps1" -CommandName run_pack
exit /b 1
