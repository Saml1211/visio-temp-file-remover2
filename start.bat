@echo off
setlocal
cd /d "%~dp0"
echo ========================================================
echo =    Visio Temp File Remover - Application Launcher    =
echo ========================================================
echo.

:: Check if Node.js is installed
where node >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: Node.js is not installed or not in PATH.
    echo Please install Node.js from https://nodejs.org/
    echo.
    pause
    exit /b 1
)

:: Check if npm is installed
where npm >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: npm is not installed or not in PATH.
    echo Please reinstall Node.js from https://nodejs.org/
    echo.
    pause
    exit /b 1
)

:: Check if package.json exists
if not exist "package.json" (
    echo ERROR: package.json not found.
    echo Please ensure you're running this from the correct directory.
    echo.
    pause
    exit /b 1
)

:: Install dependencies if node_modules doesn't exist
if not exist "node_modules\" (
    echo Installing dependencies... This may take a moment.
    call npm install
    if %ERRORLEVEL% neq 0 (
        echo ERROR: Failed to install dependencies.
        echo.
        pause
        exit /b 1
    )
    echo Dependencies installed successfully.
    echo.
)

echo Starting Visio Temp File Remover on http://localhost:3000
echo.
echo * You can access the application by opening your browser and navigating to:
echo   http://localhost:3000
echo * Press Ctrl+C to stop the server.
echo.

:: Start the application
node app.js
set "RC=%ERRORLEVEL%"

:: This will only execute if the application crashes
echo.
if "%RC%"=="3221225786" (
  echo Server stopped by user (Ctrl+C).
) else if "%RC%"=="0" (
  echo Server exited normally.
) else (
  echo ERROR: The application has stopped unexpectedly. (exit code %RC%)
)
echo.
pause
