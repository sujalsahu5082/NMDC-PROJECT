@echo off
setlocal enabledelayedexpansion

set "ROOT=%~dp0"
set "PYTHON=%ROOT%.venv\Scripts\python.exe"
set "APP=%ROOT%backend\app.py"
set "REQUIREMENTS=%ROOT%backend\requirements.txt"
set "URL=http://127.0.0.1:5000/"

echo =====================================
echo NMDC Dashboard - Startup Script
echo =====================================

if not exist "%APP%" (
  echo Backend app not found at "%APP%"
  exit /b 1
)

REM Check if virtual environment exists
if not exist "%ROOT%.venv" (
  echo Creating virtual environment...
  python -m venv "%ROOT%.venv"
  if errorlevel 1 (
    echo Failed to create virtual environment
    exit /b 1
  )
  echo Virtual environment created successfully.
)

REM Install requirements
echo Installing dependencies...
"!PYTHON!" -m pip install -q --upgrade pip
"!PYTHON!" -m pip install -q -r "!REQUIREMENTS!"
if errorlevel 1 (
  echo Failed to install dependencies
  exit /b 1
)
echo Dependencies installed.

echo.
echo Starting NMDC backend...
start "NMDC Backend" cmd /k ""!PYTHON!" "!APP!""

echo Waiting for the server to come up...
timeout /t 3 /nobreak >nul

echo Opening dashboard in browser...
start "" "%URL%"

echo.
echo Dashboard is running at %URL%
echo Press Ctrl+C in the backend window to stop the server.

endlocal