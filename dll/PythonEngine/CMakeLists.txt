add_library(PythonEngine SHARED PythonEngine.cpp)
target_compile_definitions(PythonEngine PRIVATE PYTHONENGINE_EXPORTS)
target_include_directories(PythonEngine PUBLIC
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/PythonEngine
    ${CMAKE_CURRENT_SOURCE_DIR}
)
target_link_libraries(PythonEngine PRIVATE Qt6::Core Qt6::Network)
set_target_properties(PythonEngine PROPERTIES OUTPUT_NAME "PythonEngine" PREFIX "")
