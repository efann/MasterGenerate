@echo off

@REM Get path of script
set SCRIPT_PATH=%~dp0

echo cd %SCRIPT_PATH%
cd "%SCRIPT_PATH%"

"%PROGRAMFILES%\LibreOffice\program\python.exe" main.py

pause