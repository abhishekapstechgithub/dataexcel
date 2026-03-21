@echo off
:: ============================================================
::  QtSpreadsheet — Windows Build + Installer Script
::  Requirements: Qt 6.5+, CMake 3.25+, Visual Studio 2019/2022
::                NSIS (for installer)
:: ============================================================

set BUILD_TYPE=Release
set QT_DIR=C:\Qt\6.5.3\msvc2019_64
set BUILD_DIR=build-windows
set DEPLOY_DIR=deploy

echo.
echo ==================================================
echo  QtSpreadsheet Windows Builder
echo  Build type : %BUILD_TYPE%
echo  Qt path    : %QT_DIR%
echo ==================================================
echo.

:: ── Step 1: Configure ────────────────────────────────────────────────────
echo [1/4] Configuring...
cmake -B %BUILD_DIR% ^
    -DCMAKE_BUILD_TYPE=%BUILD_TYPE% ^
    -DCMAKE_PREFIX_PATH="%QT_DIR%"
if errorlevel 1 goto :error

:: ── Step 2: Build ────────────────────────────────────────────────────────
echo [2/4] Building...
cmake --build %BUILD_DIR% --config %BUILD_TYPE% --parallel
if errorlevel 1 goto :error

:: ── Step 3: Deploy Qt DLLs ───────────────────────────────────────────────
echo [3/4] Deploying Qt dependencies...
if exist %DEPLOY_DIR% rmdir /s /q %DEPLOY_DIR%
mkdir %DEPLOY_DIR%
copy %BUILD_DIR%\bin\QtSpreadsheet.exe %DEPLOY_DIR%\
copy %BUILD_DIR%\bin\*.dll %DEPLOY_DIR%\ 2>nul
"%QT_DIR%\bin\windeployqt.exe" --release --no-translations %DEPLOY_DIR%\QtSpreadsheet.exe
if errorlevel 1 goto :error

:: ── Step 4: Build NSIS Installer ─────────────────────────────────────────
echo [4/4] Building installer...
copy installer\QtSpreadsheet.nsi %DEPLOY_DIR%\
cd %DEPLOY_DIR%
makensis QtSpreadsheet.nsi
if errorlevel 1 (cd .. && goto :error)
cd ..

echo.
echo ==================================================
echo  BUILD SUCCESSFUL
echo  Installer: %DEPLOY_DIR%\QtSpreadsheet-Setup.exe
echo ==================================================
goto :end

:error
echo.
echo !! BUILD FAILED !!
exit /b 1

:end
