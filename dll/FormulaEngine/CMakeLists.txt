set(SOURCES
    FormulaEngine.cpp
    Tokenizer.cpp
    Parser.cpp
    Functions.cpp
)

add_library(FormulaEngine SHARED ${SOURCES})
target_compile_definitions(FormulaEngine PRIVATE FORMULAENGINE_EXPORTS)
target_include_directories(FormulaEngine PUBLIC
    ${CMAKE_SOURCE_DIR}/include/FormulaEngine
    ${CMAKE_SOURCE_DIR}/include
)
target_link_libraries(FormulaEngine PRIVATE Qt6::Core)
set_target_properties(FormulaEngine PROPERTIES
    OUTPUT_NAME "FormulaEngine"
    PREFIX ""
)
