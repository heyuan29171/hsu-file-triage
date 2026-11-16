@echo off
chcp 65001 >nul
setlocal
where pwsh >nul 2>nul
if errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0hsu.ps1" -PinConsole %*
) else (
    pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0hsu.ps1" -PinConsole %*
)
echo.
pause
