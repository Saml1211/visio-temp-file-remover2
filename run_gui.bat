@echo off
setlocal
cd /d "%~dp0"
title Visio Temp File Remover
echo Starting Visio Temp File Remover...
rem Prefer Windows Python launcher, fallback to python
where py >nul 2>nul
if %ERRORLEVEL%==0 (
  py -3 "%~dp0\visio_gui.py"
) else (
  where python >nul 2>nul
  if %ERRORLEVEL%==0 (
    python "%~dp0\visio_gui.py"
  ) else (
    echo ERROR: Python 3 is not installed or not in PATH.
    pause
    exit /b 1
  )
)
echo.
pause