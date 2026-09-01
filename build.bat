@echo off
REM Builds the standalone EDMSDataBridge.exe.
REM Run from anywhere - this activates the project's .venv itself if one
REM exists next to this script. If there's no .venv yet, first run:
REM   python -m venv .venv
REM   .venv\Scripts\activate
REM   pip install -r requirements-dev.txt

if exist "%~dp0.venv\Scripts\activate.bat" (
    call "%~dp0.venv\Scripts\activate.bat"
)

pyinstaller --onefile --windowed --name "EDMSDataBridge" --version-file version_info.txt --icon assets\logo.ico --add-data "assets;assets" edms_databridge.py

if errorlevel 1 (
    echo.
    echo Build FAILED - see the error above. Common cause: dependencies
    echo aren't installed. From this folder, run:
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements-dev.txt
    pause
    exit /b 1
)

echo.
echo Build complete. The exe is in the "dist" folder:
echo   dist\EDMSDataBridge.exe
pause
