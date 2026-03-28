@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul && set "PS_EXE=pwsh"
set "NO_PAUSE_FLAG=%TEMP%\excel2llm_nopause_%RANDOM%_%RANDOM%.flag"
if exist "%NO_PAUSE_FLAG%" del /f /q "%NO_PAUSE_FLAG%" >nul 2>nul
set "EXCEL2LLM_NO_PAUSE_FLAG=%NO_PAUSE_FLAG%"

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "& '%ROOT_DIR%scripts\invoke_excel2llm.ps1' @args" -- %*
set "EXIT_CODE=%ERRORLEVEL%"

if /I not "%EXCEL2LLM_NO_PAUSE%"=="1" if not exist "%NO_PAUSE_FLAG%" pause
if exist "%NO_PAUSE_FLAG%" del /f /q "%NO_PAUSE_FLAG%" >nul 2>nul

exit /b %EXIT_CODE%
