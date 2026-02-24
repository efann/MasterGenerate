@echo off

@REM Get path of script
set SCRIPT_PATH=%~dp0
set PYTHON=%PROGRAMFILES%\LibreOffice\program\python.exe

echo Path of python interpreter: %PYTHON%
REM echo. is just a blank line.
echo.

if not exist "%PYTHON%" (
    echo %PYTHON% cannot be found.
    echo Ensure that LibreOffice has been installed.
    echo.
    pause
    exit
)

echo cd %SCRIPT_PATH%
cd "%SCRIPT_PATH%"

REM You need the quotes around PYTHON as it contains spaces.
"%PYTHON%" main.py

pause