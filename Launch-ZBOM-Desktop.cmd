@echo off
REM ZBOM Desktop Launcher — Last updated: 8/29/2026
setlocal
cd /d "%~dp0desktop-client"

if not exist "node_modules\electron\dist\electron.exe" (
  echo ZBOM preview runtime was not found.
  pause
  exit /b 1
)

start "ZBOM Desktop Application" "node_modules\electron\dist\electron.exe" .
endlocal
