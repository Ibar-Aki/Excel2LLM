@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul && set "PS_EXE=pwsh"
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\extract_excel.ps1" %*
