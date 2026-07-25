@echo off
setlocal
cd /d "%~dp0"
title EnergyPlus Review Packager v1.03

where py >nul 2>nul
if errorlevel 1 (
  echo Python 3.10+ was not found. Install it from python.org.
  pause
  exit /b 2
)

py -3 -c "import bs4,reportlab" >nul 2>nul
if errorlevel 1 (
  echo First run: installing Beautiful Soup and ReportLab...
  py -3 -m pip install --user "beautifulsoup4>=4.12,<5" "reportlab>=4,<5"
  if errorlevel 1 (
    echo Dependency installation failed.
    pause
    exit /b 3
  )
)

start "" pyw -3 "%~dp0EnergyPlusReviewPackager.py"
exit /b 0
