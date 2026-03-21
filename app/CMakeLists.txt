set(SOURCES
    main.cpp
    MainWindow.cpp
    SpreadsheetView.cpp
    FormatCellsDialog.cpp
    FindReplaceDialog.cpp
)

add_executable(QtSpreadsheet ${SOURCES})
target_include_directories(QtSpreadsheet PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetEngine
    ${CMAKE_SOURCE_DIR}/include/FormulaEngine
    ${CMAKE_SOURCE_DIR}/include/FileLoader
    ${CMAKE_SOURCE_DIR}/include/RibbonUI
    ${CMAKE_CURRENT_SOURCE_DIR}
)
target_link_libraries(QtSpreadsheet PRIVATE
    Qt6::Core Qt6::Gui Qt6::Widgets Qt6::Concurrent
    SpreadsheetCore SpreadsheetEngine FormulaEngine FileLoader RibbonUI
)
set_target_properties(QtSpreadsheet PROPERTIES
    WIN32_EXECUTABLE TRUE
    OUTPUT_NAME "QtSpreadsheet"
)
