@echo off
echo ========================================================
echo Building HI-LEX Sub-Process Traceability Executable
echo ========================================================
echo.

:: Check if PyInstaller is installed
python3-64.exe -m pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller is not installed. Installing now...
    python3-64.exe -m pip install pyinstaller
)

echo.
echo Compiling... This may take a minute or two...
echo.

:: Run PyInstaller
python3-64.exe -m PyInstaller --noconsole --onefile --distpath "..\dist" --workpath "..\build" --specpath ".." --add-data "..\assets;assets" --icon="..\assets\logo_en.png" --version-file=file_version_info.txt "..\src\main.py"

echo.
echo ========================================================
echo Build Complete!
echo You can find your Sub-Process Traceability.exe in the 'dist' folder.
echo ========================================================
pause
