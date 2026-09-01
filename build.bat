@echo off
REM Builds the standalone EDMSDataBridge.exe
REM Run this from the project folder after activating your venv / installing requirements.txt

pyinstaller --onefile --windowed --name "EDMSDataBridge" --version-file version_info.txt --icon assets\logo.ico --add-data "assets;assets" edms_databridge.py

echo.
echo Build complete. The exe is in the "dist" folder:
echo   dist\EDMSDataBridge.exe
pause
