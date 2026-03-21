@echo off
REM ============================================================
REM  Build the QtSpreadsheet NSIS installer
REM  Run this AFTER build.bat has produced build-win\bin\
REM  Requirements: NSIS 3.x installed
REM ============================================================
setlocal

set NSIS="C:\Program Files (x86)\NSIS\makensis.exe"
if not exist %NSIS% set NSIS="C:\Program Files\NSIS\makensis.exe"
if not exist %NSIS% (
    echo ERROR: NSIS not found. Download from https://nsis.sourceforge.io/
    exit /b 1
)

echo [1/2] Running windeployqt to collect Qt DLLs...
set QT_DIR=C:\Qt\6.5.3\msvc2022_64
"%QT_DIR%\bin\windeployqt.exe" --release --no-translations --no-opengl-sw ..\build-win\bin\QtSpreadsheet.exe

echo [2/2] Building installer with NSIS...
%NSIS% QtSpreadsheet.nsi

if %errorlevel% == 0 (
    echo.
    echo  ===================================================
    echo   Installer created: QtSpreadsheet-1.0.0-Setup.exe
    echo  ===================================================
) else (
    echo NSIS failed with code %errorlevel%
)
