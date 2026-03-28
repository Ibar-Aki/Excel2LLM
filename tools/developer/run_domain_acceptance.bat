@echo off
setlocal

for %%I in ("%~dp0..\..") do set "ROOT_DIR=%%~fI\"

where pwsh >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%scripts\run_domain_acceptance.ps1" %*
) else (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%scripts\run_domain_acceptance.ps1" %*
)
