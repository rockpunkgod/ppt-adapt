@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0officecli-ppt-design\scripts\rebuild_demo.ps1" %*
if errorlevel 1 (
 echo Rebuild failed. Please share the message above with your Agent.
 pause
 exit /b 1
)
echo Done. The output folder contains the editable PPTX and previews.
pause
