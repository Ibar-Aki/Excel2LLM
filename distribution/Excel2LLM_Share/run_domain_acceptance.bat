@echo off
setlocal

set SCRIPT_DIR=%~dp0

where pwsh >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\run_domain_acceptance.ps1" %*
) else (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\run_domain_acceptance.ps1" %*
)
