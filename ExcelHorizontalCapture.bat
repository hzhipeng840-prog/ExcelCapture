@echo off
setlocal
cd /d "%~dp0"
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File "%~dp0ExcelHorizontalCapture.ps1"
if errorlevel 1 (
  echo.
  echo Tool exited with an error.
  pause
)
