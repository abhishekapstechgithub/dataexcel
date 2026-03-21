# SpreadsheetTableModel.h has Q_OBJECT and lives in include/ — explicitly MOC it.
qt_wrap_cpp(SPREADSHEETMODEL_MOC
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetEngine/SpreadsheetTableModel.h
)

set(SOURCES SpreadsheetTableModel.cpp ${SPREADSHEETMODEL_MOC})
add_library(SpreadsheetEngine SHARED ${SOURCES})
target_compile_definitions(SpreadsheetEngine PRIVATE SPREADSHEETENGINE_EXPORTS)
target_include_directories(SpreadsheetEngine PUBLIC
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetEngine
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
)
target_link_libraries(SpreadsheetEngine PRIVATE Qt6::Core Qt6::Widgets SpreadsheetCore)
set_target_properties(SpreadsheetEngine PROPERTIES OUTPUT_NAME "SpreadsheetEngine" PREFIX "")
