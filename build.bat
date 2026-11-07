@echo off
setlocal
cd /d "%~dp0"

set PYEXE=.venv\Scripts\python.exe

if not exist ".venv" (
    echo [setup] creating virtual environment ...
    py -3 -m venv .venv
    if errorlevel 1 goto :err
    "%PYEXE%" -m pip install --upgrade pip
    "%PYEXE%" -m pip install -r requirements.txt
    if errorlevel 1 goto :err
)

echo [build] installing pyinstaller ...
"%PYEXE%" -m pip install --upgrade pyinstaller
if errorlevel 1 goto :err

echo [build] packaging ...
"%PYEXE%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --windowed ^
    --name TableFiller ^
    --collect-all rapidocr_onnxruntime ^
    --add-data "app;app" ^
    main.py
if errorlevel 1 goto :err

echo.
echo [done] dist\TableFiller.exe
exit /b 0

:err
echo [error] build failed.
exit /b 1
