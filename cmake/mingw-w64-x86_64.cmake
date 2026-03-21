# CMake toolchain file for cross-compiling to Windows x64 using MinGW-w64
# Usage: cmake .. -DCMAKE_TOOLCHAIN_FILE=cmake/mingw-w64-x86_64.cmake \
#                 -DCMAKE_PREFIX_PATH=/path/to/Qt/6.5.3/mingw_64

set(CMAKE_SYSTEM_NAME    Windows)
set(CMAKE_SYSTEM_VERSION 10)
set(CMAKE_SYSTEM_PROCESSOR x86_64)

# Cross-compiler — adjust prefix if yours differs
set(CROSS x86_64-w64-mingw32)
find_program(CMAKE_C_COMPILER   NAMES ${CROSS}-gcc   REQUIRED)
find_program(CMAKE_CXX_COMPILER NAMES ${CROSS}-g++   REQUIRED)
find_program(CMAKE_RC_COMPILER  NAMES ${CROSS}-windres)

set(CMAKE_FIND_ROOT_PATH_MODE_PROGRAM NEVER)
set(CMAKE_FIND_ROOT_PATH_MODE_LIBRARY ONLY)
set(CMAKE_FIND_ROOT_PATH_MODE_INCLUDE ONLY)
set(CMAKE_FIND_ROOT_PATH_MODE_PACKAGE ONLY)
