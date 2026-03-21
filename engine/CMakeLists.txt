# ── SpreadsheetEngine library ─────────────────────────────────────────────────
# Self-contained engine: sparse cell storage, formula evaluation,
# dependency graph, undo/redo. No Qt UI dependencies.

set(ENGINE_SOURCES
    src/SpreadsheetEngine.cpp
    src/EngineModel.cpp
)

set(ENGINE_HEADERS
    include/engine/CellAddress.h
    include/engine/SpreadsheetEngine.h
    include/engine/EngineModel.h
)

add_library(SpreadsheetEngineLib STATIC ${ENGINE_SOURCES} ${ENGINE_HEADERS})

target_include_directories(SpreadsheetEngineLib PUBLIC
    ${CMAKE_CURRENT_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/include/FormulaEngine
    ${CMAKE_SOURCE_DIR}/dll/FormulaEngine
)

target_link_libraries(SpreadsheetEngineLib PUBLIC
    Qt6::Core Qt6::Gui FormulaEngine
)

set_target_properties(SpreadsheetEngineLib PROPERTIES
    CXX_STANDARD 20
    CXX_STANDARD_REQUIRED ON
)
