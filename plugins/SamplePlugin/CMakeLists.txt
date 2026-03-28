add_library(SamplePlugin SHARED SamplePlugin.cpp)
target_compile_definitions(SamplePlugin PRIVATE PLUGIN_EXPORTS)
target_include_directories(SamplePlugin PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/PluginSystem
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
)
target_link_libraries(SamplePlugin PRIVATE Qt6::Core)
set_target_properties(SamplePlugin PROPERTIES OUTPUT_NAME "SamplePlugin" PREFIX "")
