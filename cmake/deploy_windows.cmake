# Deploy Qt DLLs alongside the application using windeployqt
# Usage: cmake -P cmake/deploy_windows.cmake
# Requires CMAKE_BINARY_DIR and Qt6_DIR to be set in cache

find_program(WINDEPLOYQT windeployqt
    HINTS "${Qt6_DIR}/../../../bin"
          "$ENV{QTDIR}/bin")

if(NOT WINDEPLOYQT)
    message(FATAL_ERROR "windeployqt not found. Set Qt6_DIR or QTDIR.")
endif()

set(EXE "${CMAKE_BINARY_DIR}/bin/QtSpreadsheet.exe")
execute_process(
    COMMAND "${WINDEPLOYQT}"
        --release
        --no-translations
        --no-opengl-sw
        --dir "${CMAKE_BINARY_DIR}/bin"
        "${EXE}"
    RESULT_VARIABLE ret)

if(NOT ret EQUAL 0)
    message(FATAL_ERROR "windeployqt failed (exit code ${ret})")
else()
    message(STATUS "windeployqt succeeded")
endif()
