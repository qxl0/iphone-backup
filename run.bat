@echo off
echo ============================================
echo  iPhone Photo Downloader
echo ============================================
echo.

REM Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Install dependencies if not already present
python -c "import win32com" >nul 2>&1
if errorlevel 1 (
    echo Installing required package: pywin32 ...
    pip install pywin32
    echo.
)

REM Default destination: Pictures\iPhone Photos
REM To use a custom folder, edit the line below:
set DEST=%USERPROFILE%\Pictures\iPhone Photos

echo Photos will be saved to: %DEST%
echo.
echo Make sure your iPhone is:
echo   1. Connected via USB
echo   2. Unlocked
echo   3. You tapped "Trust This Computer" on the iPhone
echo.
pause

python "%~dp0iphone_downloader.py" "%DEST%"

echo.
pause
