set(SOURCES
    main.cpp
    MainWindow.cpp
    SpreadsheetView.cpp
    FormulaBar.cpp
    SheetTabBar.cpp
    TileCache.cpp
    FormatCellsDialog.cpp
    FindReplaceDialog.cpp
)

add_executable(QtSpreadsheet WIN32 ${SOURCES})

set_target_properties(QtSpreadsheet PROPERTIES
    OUTPUT_NAME "OpenSheet"
    WIN32_EXECUTABLE TRUE
)

target_include_directories(QtSpreadsheet PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetEngine
    ${CMAKE_SOURCE_DIR}/include/FormulaEngine
    ${CMAKE_SOURCE_DIR}/include/FileLoader
    ${CMAKE_SOURCE_DIR}/include/RibbonUI
    ${CMAKE_SOURCE_DIR}/include/PluginSystem
    ${CMAKE_SOURCE_DIR}/dll/PluginSystem
    ${CMAKE_CURRENT_SOURCE_DIR}
)

target_link_libraries(QtSpreadsheet PRIVATE
    Qt6::Core
    Qt6::Gui
    Qt6::Widgets
    Qt6::Concurrent
    Qt6::Sql
    Qt6::Charts
    SpreadsheetCore
    SpreadsheetEngine
    FormulaEngine
    FileLoader
    RibbonUI
    PluginSystem
)

# Windows app icon
if(WIN32 AND EXISTS "${CMAKE_SOURCE_DIR}/resources/app.rc")
    target_sources(QtSpreadsheet PRIVATE "${CMAKE_SOURCE_DIR}/resources/app.rc")
endif()
