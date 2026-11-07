@echo off
setlocal
cd /d "%~dp0"

set PYEXE=.venv\Scripts\python.exe

if not exist ".venv" (
    echo [setup] creating virtual environment ...
    py -3 -m venv .venv
    if errorlevel 1 goto :err
    echo [setup] installing dependencies ...
    "%PYEXE%" -m pip install --upgrade pip
    "%PYEXE%" -m pip install -r requirements.txt
    if errorlevel 1 goto :err
)

echo [run] starting GUI ...
"%PYEXE%" main.py %*
goto :eof

:err
echo.
echo [error] setup failed. Please run manually:
echo   py -3 -m venv .venv
echo   .venv\Scripts\activate
echo   pip install -r requirements.txt
exit /b 1
