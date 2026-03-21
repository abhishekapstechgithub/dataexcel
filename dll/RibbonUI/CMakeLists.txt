qt_wrap_cpp(RIBBONWIDGET_MOC
    ${CMAKE_SOURCE_DIR}/include/RibbonUI/RibbonWidget.h
)
set(SOURCES RibbonWidget.cpp ColorButton.cpp ${RIBBONWIDGET_MOC})
add_library(RibbonUI SHARED ${SOURCES})
target_compile_definitions(RibbonUI PRIVATE RIBBONUI_EXPORTS)
target_include_directories(RibbonUI PUBLIC
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/RibbonUI
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
)
target_link_libraries(RibbonUI PRIVATE Qt6::Core Qt6::Widgets Qt6::Gui)
set_target_properties(RibbonUI PROPERTIES OUTPUT_NAME "RibbonUI" PREFIX "")
