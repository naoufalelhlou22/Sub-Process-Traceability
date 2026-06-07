@echo off
cd /d "%~dp0"
echo ========================================================
echo Building HI-LEX Sub-Process Traceability Executable
echo ========================================================
echo.

:: Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller is not installed. Installing now...
    python -m pip install pyinstaller
)

echo.
echo Compiling... This may take a minute or two...
echo.

:: Clean previous builds to avoid cache issues
if exist build rmdir /s /q build
if exist "Sub-Process Traceability.spec" del /f /q "Sub-Process Traceability.spec"

:: Run PyInstaller
python -m PyInstaller --noconsole --onefile --clean --name "Sub-Process Traceability" --distpath "dist" --workpath "build" --add-data "assets;assets" --icon="assets\logo_en.png" --version-file="scripts\file_version_info.txt" "src\main.py"

echo.
echo ========================================================
echo Build Complete!
echo You can find your Sub-Process Traceability.exe in the 'dist' folder.
echo ========================================================
pause
