@echo off
chcp 65001 > nul
echo Building iPhone Backup.exe...
echo.

pip install pyinstaller colorama pywin32 --quiet

pyinstaller ^
  --onefile ^
  --console ^
  --name "iPhone Backup" ^
  --hidden-import win32com.client ^
  --hidden-import win32com.shell ^
  --hidden-import pythoncom ^
  iphone_downloader.py

echo.
if exist "dist\iPhone Backup.exe" (
    echo Done! Your file is at:
    echo   dist\iPhone Backup.exe
) else (
    echo Build failed. Check the output above for errors.
)
echo.
pause
