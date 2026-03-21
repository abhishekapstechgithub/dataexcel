# ISpreadsheetCore.h has Q_OBJECT but lives in include/ — AUTOMOC won't find it.
# Use qt_wrap_cpp() to explicitly run MOC on it.
qt_wrap_cpp(ISPREADSHEETCORE_MOC
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore/ISpreadsheetCore.h
)

set(SOURCES SpreadsheetCore.cpp UndoStack.cpp ${ISPREADSHEETCORE_MOC})
add_library(SpreadsheetCore SHARED ${SOURCES})
target_compile_definitions(SpreadsheetCore PRIVATE SPREADSHEETCORE_EXPORTS)
target_include_directories(SpreadsheetCore PUBLIC
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
    ${CMAKE_SOURCE_DIR}/include/FormulaEngine
)
target_link_libraries(SpreadsheetCore PRIVATE Qt6::Core Qt6::Gui FormulaEngine)
set_target_properties(SpreadsheetCore PROPERTIES OUTPUT_NAME "SpreadsheetCore" PREFIX "")
