# IFileLoader.h has Q_OBJECT but lives in include/ — explicitly MOC it.
qt_wrap_cpp(IFILELOADER_MOC
    ${CMAKE_SOURCE_DIR}/include/FileLoader/IFileLoader.h
)

set(SOURCES FileLoader.cpp CsvLoader.cpp XlsxLoader.cpp XlsxSaxLoader.cpp ${IFILELOADER_MOC})
add_library(FileLoader SHARED ${SOURCES})
target_compile_definitions(FileLoader PRIVATE FILELOADER_EXPORTS)
target_include_directories(FileLoader PUBLIC
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/FileLoader
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
)
target_link_libraries(FileLoader PRIVATE Qt6::Core Qt6::Gui Qt6::Xml SpreadsheetCore)
set_target_properties(FileLoader PROPERTIES OUTPUT_NAME "FileLoader" PREFIX "")
