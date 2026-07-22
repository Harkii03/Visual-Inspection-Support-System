@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" goto :environment_ready

where py >nul 2>nul
if %errorlevel%==0 set "PYTHON_CMD=py -3"
if defined PYTHON_CMD goto :create_environment

where python >nul 2>nul
if %errorlevel%==0 set "PYTHON_CMD=python"
if defined PYTHON_CMD goto :create_environment

echo Python 3 is not installed.
echo Install Python 3, then run this file again.
pause
exit /b 1

:create_environment
echo Creating the application environment...
%PYTHON_CMD% -m venv .venv
if errorlevel 1 goto :error

:environment_ready

".venv\Scripts\python.exe" -c "import flask, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing required packages...
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :error
)

echo Starting Animal Image Review Assistant...
echo Close this window to stop the application.
".venv\Scripts\python.exe" app.py
exit /b 0

:error
echo.
echo Setup failed. Check your network connection and Python installation.
pause
exit /b 1
