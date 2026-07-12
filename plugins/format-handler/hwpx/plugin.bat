@echo off
REM OfficeCLI HWPX Plugin Wrapper
REM This script runs the Python plugin

setlocal

set PYTHON_SCRIPT=%~dp0plugin.py
set PYTHON=python

if not exist "%PYTHON_SCRIPT%" (
    echo Error: plugin.py not found at %PYTHON_SCRIPT%
    exit /b 1
)

REM Pass all arguments to the Python script
"%PYTHON%" "%PYTHON_SCRIPT%" %*

endlocal
exit /b %ERRORLEVEL%
