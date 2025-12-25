@echo off
REM CodeExtractPro v1.0 - Build Script for Windows
REM Creates both Release and Debug executables

echo ============================================
echo  CodeExtractPro v1.0 - Build Script
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH
    pause
    exit /b 1
)

REM Install/upgrade build tools
echo [1/5] Installing build dependencies...
pip install --upgrade pyinstaller customtkinter oletools pywin32 >nul 2>&1

REM Clean previous builds
echo [2/5] Cleaning previous builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build Release version (no console)
echo [3/5] Building Release version...
pyinstaller --clean --noconfirm build_release.spec
if errorlevel 1 (
    echo ERROR: Release build failed
    pause
    exit /b 1
)

REM Build Debug version (with console)
echo [4/5] Building Debug version...
pyinstaller --clean --noconfirm build_debug.spec
if errorlevel 1 (
    echo ERROR: Debug build failed
    pause
    exit /b 1
)

REM Cleanup
echo [5/5] Cleaning up...
if exist "build" rmdir /s /q "build"
del /q *.spec.bak 2>nul

echo.
echo ============================================
echo  Build completed successfully!
echo ============================================
echo.
echo Output files:
echo   - dist\CodeExtractPro.exe (Release)
echo   - dist\CodeExtractPro_Debug.exe (Debug)
echo.
pause
