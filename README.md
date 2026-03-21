# OpenSheet — Full-Featured Spreadsheet Engine

## Architecture Overview

```
┌────────────────────────────────────────────────────────────┐
│                    UI Layer (app/)                          │
│  MainWindow → RibbonWidget → SpreadsheetView → SheetBar    │
└────────────────────┬───────────────────────────────────────┘
                     │  ISpreadsheetCore (interface)
┌────────────────────▼───────────────────────────────────────┐
│              SpreadsheetCore.dll (adapter)                  │
│   Bridges ISpreadsheetCore interface to SpreadsheetEngine   │
└────────────────────┬───────────────────────────────────────┘
                     │
┌────────────────────▼───────────────────────────────────────┐
│           SpreadsheetEngineLib (static library)             │
│  ┌─────────────────┐  ┌─────────────────────────────────┐  │
│  │ SpreadsheetEngine│  │ EngineModel (QAbstractTableModel)│  │
│  │                  │  │ - Virtual 1M×16K grid           │  │
│  │ Sparse storage:  │  │ - data() only called for        │  │
│  │ unordered_map    │  │   VISIBLE cells (lazy load)     │  │
│  │ <CellAddress,    │  └─────────────────────────────────┘  │
│  │  EngineCell>     │                                        │
│  │                  │  ┌─────────────────────────────────┐  │
│  │ Dependency graph │  │ CellAddress                     │  │
│  │ auto-recalc      │  │ - 64-bit packed key             │  │
│  │ topological sort │  │ - O(1) hash                     │  │
│  └─────────────────┘  │ - Supports 1,048,576 × 16,384   │  │
│                        └─────────────────────────────────┘  │
└────────────────────┬───────────────────────────────────────┘
                     │
┌────────────────────▼───────────────────────────────────────┐
│              FormulaEngine.dll                              │
│  80+ functions: SUM, AVERAGE, IF, VLOOKUP, PMT, STDEV...   │
│  Tokenizer → Parser → FunctionRegistry → Result            │
└────────────────────────────────────────────────────────────┘
```

## Key Technical Decisions

### 1. Sparse Storage
```cpp
std::unordered_map<CellAddress, EngineCell> cells;
```
- Only non-empty cells consume memory
- 1M×16K grid with 1000 filled cells = 1000 hash entries
- O(1) read/write with 64-bit packed key

### 2. Virtual Grid (Lazy Loading)
```cpp
// EngineModel only gets data() called for VISIBLE cells
int rowCount() { return qMax(200, usedRows + 50); }  // virtual max
QVariant data(idx, role) {
    if (!engine->cellExists(row, col)) return {};  // fast path
    // only loads when viewport requests it
}
```

### 3. Dependency Graph (Auto-Recalc)
```cpp
// When A1 changes → find all cells that use A1 → recalc them
void recalcDependents(sheet, changedAddr) {
    BFS through dependents graph
    → topological order evaluation
    → only recalcs cells that actually need it
}
```

### 4. Formula Engine (80+ functions)
- **Math**: SUM, AVERAGE, MIN, MAX, ROUND, ABS, POWER, SQRT, MOD, INT, CEILING, FLOOR
- **Trig**: SIN, COS, TAN, ASIN, ACOS, ATAN, PI
- **Statistical**: STDEV, VAR, MEDIAN, LARGE, SMALL, RANK, COUNTIF, SUMIF
- **Text**: LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, CONCATENATE, SUBSTITUTE, FIND, TEXT
- **Logical**: IF, AND, OR, NOT, IFERROR, IFS, SWITCH
- **Date**: NOW, TODAY, YEAR, MONTH, DAY, DATE, DATEDIF, EDATE, DAYS, NETWORKDAYS
- **Lookup**: INDEX, MATCH, CHOOSE
- **Financial**: PMT, PV, FV, NPV

## Build Instructions

### Prerequisites
- Qt 6.5+
- CMake 3.22+
- GCC 11+ or MSVC 2019+

### Linux/macOS
```bash
bash build.sh
```

### Windows
```bat
build.bat
```

### Manual
```bash
cmake -B build -DCMAKE_BUILD_TYPE=Release -DCMAKE_PREFIX_PATH=/path/to/Qt/6.5.3/gcc_64
cmake --build build --config Release --parallel
```

## Module Structure
```
OpenSheet/
├── engine/                    # Core engine (static lib)
│   ├── include/engine/
│   │   ├── CellAddress.h      # 64-bit packed cell coordinate
│   │   ├── SpreadsheetEngine.h# Main engine API
│   │   └── EngineModel.h      # Virtual QAbstractTableModel
│   └── src/
│       ├── SpreadsheetEngine.cpp  # Sparse storage + dep graph
│       └── EngineModel.cpp        # Lazy-loading model
│
├── dll/
│   ├── FormulaEngine/         # Formula parser + 80+ functions
│   ├── SpreadsheetCore/       # ISpreadsheetCore adapter
│   ├── SpreadsheetEngine/     # Qt model (legacy, wraps engine)
│   ├── FileLoader/            # CSV + XLSX streaming loader
│   └── RibbonUI/              # Ribbon widget
│
├── include/                   # Public interfaces
├── app/                       # Application shell (MainWindow)
└── installer/                 # NSIS installer script
```
