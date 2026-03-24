@echo off
:: ═══════════════════════════════════════════════════════════════════════════════
::  build.bat — Quick build script for Windows (requires Qt6 and CMake in PATH)
:: ═══════════════════════════════════════════════════════════════════════════════
setlocal

:: Detect Qt6 location from common paths
if exist "%Qt6_DIR%" (
    set "QT_PREFIX=%Qt6_DIR%"
) else if exist "C:\Qt\6.6.0\msvc2019_64" (
    set "QT_PREFIX=C:\Qt\6.6.0\msvc2019_64"
) else if exist "C:\Qt\6.5.3\msvc2019_64" (
    set "QT_PREFIX=C:\Qt\6.5.3\msvc2019_64"
) else (
    echo ERROR: Qt6 not found. Set Qt6_DIR environment variable.
    exit /b 1
)

echo Building OpenSheet with Qt at: %QT_PREFIX%
echo.

:: Configure
cmake -B build -S . ^
    -DCMAKE_BUILD_TYPE=Release ^
    -DCMAKE_PREFIX_PATH="%QT_PREFIX%" ^
    -G "Visual Studio 17 2022" -A x64

if errorlevel 1 (
    echo CMake configure FAILED
    exit /b 1
)

:: Build
cmake --build build --config Release --parallel %NUMBER_OF_PROCESSORS%

if errorlevel 1 (
    echo Build FAILED
    exit /b 1
)

echo.
echo Build succeeded! Executable: build\bin\OpenSheet.exe
echo.

:: Deploy Qt DLLs
if exist "build\bin\OpenSheet.exe" (
    echo Deploying Qt DLLs...
    "%QT_PREFIX%\bin\windeployqt.exe" --release --no-translations ^
        --no-system-d3d-compiler "build\bin\OpenSheet.exe"
    echo Done!
)

endlocal
