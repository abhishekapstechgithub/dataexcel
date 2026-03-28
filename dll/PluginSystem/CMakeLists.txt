add_library(PluginSystem SHARED PluginManager.cpp)
target_compile_definitions(PluginSystem PRIVATE PLUGINSYSTEM_EXPORTS)
target_include_directories(PluginSystem PUBLIC
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/PluginSystem
    ${CMAKE_SOURCE_DIR}/include/SpreadsheetCore
    ${CMAKE_CURRENT_SOURCE_DIR}
)
target_link_libraries(PluginSystem PRIVATE Qt6::Core SpreadsheetCore)
set_target_properties(PluginSystem PROPERTIES OUTPUT_NAME "PluginSystem" PREFIX "")
