@echo off
setlocal
set SCRIPT_DIR=%~dp0
start "" powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%context_menu_tool.ps1" -Gui
