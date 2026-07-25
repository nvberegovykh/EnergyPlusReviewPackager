@echo off
setlocal
cd /d "%~dp0"
title Build EnergyPlus Review Packager v1.03

where py >nul 2>nul
if errorlevel 1 (
  echo Python Launcher was not found. Install Python 3.10+ from python.org.
  pause
  exit /b 2
)

py -3 -m pip install --upgrade pyinstaller beautifulsoup4 reportlab customtkinter tkinterdnd2
if errorlevel 1 goto :failed

py -3 -m PyInstaller --noconfirm --clean EnergyPlusReviewPackager_release.spec
if errorlevel 1 goto :failed

echo.
echo Built:
echo   dist\EnergyPlusReviewPackager_v1.03.exe
pause
exit /b 0

:failed
echo.
echo Build failed. Review the messages above.
pause
exit /b 1
