@echo off
setlocal
set "ZBOM_PREVIEW=1"
set "ZBOM_PREVIEW=1"
cd /d "%~dp0"

if not exist "node_modules\electron\dist\electron.exe" (
  echo ZBOM preview runtime was not found.
  echo Please keep this file inside the desktop-client folder.
  pause
  exit /b 1
)

start "ZBOM Preview" "node_modules\electron\dist\electron.exe" .
endlocal
