@echo off
title Building Traceability System Executables
echo ============================================================
echo  BUILDING EXECUTABLES
echo ============================================================

echo.
echo Installing PyInstaller if needed...
python -m pip install pyinstaller

echo.
echo --- Building Main Application ---
if exist "Sub-Process Traceability.spec" (
    python -m PyInstaller "Sub-Process Traceability.spec" --clean
) else (
    echo Spec file "Sub-Process Traceability.spec" not found.
)

echo.
echo --- Building Quality Application ---
if exist "Quality App.spec" (
    python -m PyInstaller "Quality App.spec" --clean
) else (
    echo Spec file "Quality App.spec" not found.
)

echo.
echo ============================================================
echo  BUILD COMPLETE
echo ============================================================
echo You can find the compiled .exe files inside the "dist" folder.
echo.
pause
Traceback (most recent call last):
  File "C:\Users\Mr_X\Desktop\Sub-Process Traceability\src\quality_app.py", line 15, in <module>
    from config import *
ModuleNotFoundError: No module named 'config'