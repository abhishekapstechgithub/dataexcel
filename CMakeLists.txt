cmake_minimum_required(VERSION 3.20)
project(OpenSheet VERSION 1.0.0 LANGUAGES CXX)

# ── C++ standard ──────────────────────────────────────────────────────────────
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# ── Qt auto-tools ─────────────────────────────────────────────────────────────
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC OFF)

# ── Output directories ────────────────────────────────────────────────────────
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)

# Multi-config generators (MSVC, Ninja Multi-Config)
foreach(cfg IN ITEMS Debug Release RelWithDebInfo MinSizeRel)
    string(TOUPPER ${cfg} CFG)
    set(CMAKE_RUNTIME_OUTPUT_DIRECTORY_${CFG} ${CMAKE_BINARY_DIR}/bin)
    set(CMAKE_LIBRARY_OUTPUT_DIRECTORY_${CFG} ${CMAKE_BINARY_DIR}/bin)
    set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY_${CFG} ${CMAKE_BINARY_DIR}/lib)
endforeach()

# ── Qt6 packages ─────────────────────────────────────────────────────────────
find_package(Qt6 6.5 REQUIRED COMPONENTS
    Core Gui Widgets Concurrent Xml Charts Sql Network)
qt_standard_project_setup()

# ── Global compiler flags ─────────────────────────────────────────────────────
if(MSVC)
    add_compile_options(/W3 /utf-8 /MP)
    add_compile_definitions(_UNICODE UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
else()
    add_compile_options(-Wall -Wextra -Wno-unused-parameter)
endif()

# ── Global include paths ──────────────────────────────────────────────────────
include_directories(
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/engine/include
)

# ── Submodules ────────────────────────────────────────────────────────────────
add_subdirectory(engine)
add_subdirectory(dll/FormulaEngine)
add_subdirectory(dll/SpreadsheetCore)
add_subdirectory(dll/SpreadsheetEngine)
add_subdirectory(dll/FileLoader)
add_subdirectory(dll/RibbonUI)
add_subdirectory(dll/PluginSystem)
add_subdirectory(dll/PythonEngine)
add_subdirectory(plugins/SamplePlugin)
add_subdirectory(app)
