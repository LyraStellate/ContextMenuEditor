@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%context_menu_tool.ps1"
if errorlevel 1 (
  echo.
  echo An error occurred.
)
echo.
pause
