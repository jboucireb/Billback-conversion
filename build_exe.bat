@echo off
title Monin Billback - Build Installer
echo ============================================================
echo  Monin Billback Processor - Build .exe
echo ============================================================
echo.

REM Try py launcher first (always works on Windows), then python
set PYCMD=
py --version >nul 2>&1
if not errorlevel 1 (
    set PYCMD=py
    goto :found_python
)
python --version >nul 2>&1
if not errorlevel 1 (
    set PYCMD=python
    goto :found_python
)
echo ERROR: Python not found.
echo Please install Python 3.12 from https://python.org
echo Make sure to check "Add Python to PATH" during install.
pause
exit /b 1

:found_python
echo Using: %PYCMD%
echo.

echo [1/3] Installing required libraries...
%PYCMD% -m pip install --upgrade pip --quiet
%PYCMD% -m pip install -r requirements.txt --quiet
%PYCMD% -m pip install pyinstaller --quiet
if errorlevel 1 (
    echo.
    echo ERROR: pip install failed.
    echo Check your internet connection or company proxy settings.
    pause
    exit /b 1
)
echo       Done.

echo [2/3] Building standalone .exe (this takes ~60 seconds)...
%PYCMD% -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --name "Monin_Billback_Processor" ^
  --hidden-import pdfplumber ^
  --hidden-import pdfminer ^
  --hidden-import pdfminer.high_level ^
  --hidden-import pdfminer.layout ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  --hidden-import pandas ^
  --hidden-import email ^
  --hidden-import email.parser ^
  --hidden-import email.policy ^
  --hidden-import cgi ^
  --hidden-import uuid ^
  billback_app.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed. See above for details.
    pause
    exit /b 1
)
echo       Done.

echo [3/3] Cleaning up build files...
rmdir /s /q build >nul 2>&1
del /q Monin_Billback_Processor.spec >nul 2>&1
echo       Done.

echo.
echo ============================================================
echo  SUCCESS!
echo  Your .exe is ready at:
echo    dist\Monin_Billback_Processor.exe
echo.
echo  Copy that .exe to any PC and double-click it.
echo  No Python or installs needed on those PCs.
echo ============================================================
echo.
pause
