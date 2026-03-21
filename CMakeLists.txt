cmake_minimum_required(VERSION 3.22)
project(QtSpreadsheet VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_INCLUDE_CURRENT_DIR ON)

# Output directories — flat bin/ for all configs (works with both GCC and MSVC)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY         ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY_RELEASE ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY_DEBUG   ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY         ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY_RELEASE ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY         ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY_RELEASE ${CMAKE_BINARY_DIR}/lib)

find_package(Qt6 REQUIRED COMPONENTS Core Gui Widgets Concurrent Xml)

include_directories(
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetEngine
    ${CMAKE_SOURCE_DIR}/include/FormulaEngine
    ${CMAKE_SOURCE_DIR}/include/FileLoader
    ${CMAKE_SOURCE_DIR}/include/RibbonUI
)

add_subdirectory(engine)
add_subdirectory(dll/FormulaEngine)
add_subdirectory(dll/SpreadsheetCore)
add_subdirectory(dll/SpreadsheetEngine)
add_subdirectory(dll/FileLoader)
add_subdirectory(dll/RibbonUI)
add_subdirectory(app)

install(TARGETS
    FormulaEngine SpreadsheetCore SpreadsheetEngine FileLoader RibbonUI QtSpreadsheet
    RUNTIME DESTINATION bin
    LIBRARY DESTINATION bin
)
# SpreadsheetEngineLib is static — no install needed
