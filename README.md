# QtSpreadsheet

High-performance spreadsheet application built with **C++20 + Qt 6**, modular **DLL architecture**,
and a fully virtual grid that handles **1,000,000+ rows** with smooth scrolling.

---

## Features

| Feature | Details |
|---------|---------|
| Ribbon UI | Excel-style Home tab: Clipboard, Font, Alignment, Number, Cells, Editing |
| Virtual grid | QTableView + QAbstractTableModel — only visible cells ever loaded |
| Formula engine | 40+ functions: SUM, AVERAGE, MIN, MAX, IF, VLOOKUP, STDEV, TEXT, … |
| File I/O | Streaming CSV and XLSX read/write |
| Undo/redo | 200-level command stack |
| Threading | QtConcurrent async file loading with live progress bar |
| DLL architecture | 5 independent DLLs, minimal code in main.exe |
| Multi-sheet | Add, rename, remove, switch sheets |
| Sort/filter | Range sort ascending/descending |
| Find & Replace | Case-sensitive, whole-sheet replacement |
| Format Cells | Font, number format, alignment, wrap text |
| Cell formatting | Bold, italic, underline, text/fill color, number formats |
| Merge cells | Merge/unmerge selection |

---

## Requirements

| Tool | Minimum version |
|------|----------------|
| C++ compiler | GCC 12, Clang 16, or MSVC 2022 (C++20) |
| CMake | 3.25 |
| Qt | 6.5.0 (Widgets, Core, Gui, Concurrent, Xml) |

### Install Qt (options)

**Official Qt installer (recommended)**
Download from https://www.qt.io/download-qt-installer and install Qt 6.5 or later.

**vcpkg**
```sh
vcpkg install qt6-base qt6-tools
```

**apt (Ubuntu 24.04)**
```sh
sudo apt install qt6-base-dev qt6-tools-dev cmake build-essential
```

**brew (macOS)**
```sh
brew install qt@6 cmake
```

---

## Build — Windows

```bat
REM Edit QT_DIR if your Qt is installed elsewhere
build.bat Release "C:\Qt\6.5.3\msvc2022_64"
```

Output: `build-win\bin\QtSpreadsheet.exe` + all DLLs in the same folder.

**Manual steps**
```bat
cmake -B build-win -G "Visual Studio 17 2022" -A x64 ^
    -DCMAKE_PREFIX_PATH="C:\Qt\6.5.3\msvc2022_64"
cmake --build build-win --config Release --parallel
```

---

## Build — Linux / macOS

```sh
chmod +x build.sh
./build.sh Release /opt/Qt/6.5.3/gcc_64
```

Output: `build-linux/bin/QtSpreadsheet`

**Manual steps**
```sh
cmake -B build-linux \
    -DCMAKE_BUILD_TYPE=Release \
    -DCMAKE_PREFIX_PATH=/opt/Qt/6.5.3/gcc_64
cmake --build build-linux --parallel $(nproc)
```

---

## Project structure

```
QtSpreadsheet/
├── CMakeLists.txt                 Root CMake — builds all targets
├── build.bat                      Windows one-click build
├── build.sh                       Linux/macOS one-click build
│
├── include/
│   ├── SpreadsheetCore/
│   │   └── ISpreadsheetCore.h     Cell, CellFormat, ISpreadsheetCore interface
│   ├── SpreadsheetEngine/
│   │   └── SpreadsheetTableModel.h Virtual Qt model interface
│   ├── FormulaEngine/
│   │   └── FormulaEngine.h        Formula evaluation public API
│   ├── FileLoader/
│   │   └── IFileLoader.h          Streaming file loader interface
│   └── RibbonUI/
│       └── RibbonWidget.h         Ribbon signals/state public API
│
├── dll/
│   ├── FormulaEngine/             Tokenizer → Parser → 40+ functions
│   ├── SpreadsheetCore/           Cell store + undo stack + formula recalc
│   ├── SpreadsheetEngine/         QAbstractTableModel (viewport-only)
│   ├── FileLoader/                CSV + XLSX streaming loader/saver
│   └── RibbonUI/                  Full Excel-style Home ribbon
│
├── app/
│   ├── main.cpp                   Entry point, QApplication
│   ├── MainWindow.{h,cpp}         Controller — wires DLLs together
│   ├── SpreadsheetView.{h,cpp}    QTableView subclass + selection logic
│   ├── FormatCellsDialog.{h,cpp}  Format Cells dialog
│   └── FindReplaceDialog.{h,cpp}  Find & Replace dialog
│
├── cmake/
│   └── deploy_windows.cmake       windeployqt helper
│
├── resources/
│   └── app.rc                     Windows version info + icon resource
│
└── data/
    ├── sample_data.csv            1,000-row sample (included)
    └── generate_sample.py         Generator: python3 generate_sample.py 1000000
```

---

## DLL architecture

Each DLL exports a single `extern "C"` factory function:

| DLL | Factory | Responsibility |
|-----|---------|---------------|
| `FormulaEngine.dll` | — (used directly) | Tokenize → parse → evaluate formulas |
| `SpreadsheetCore.dll` | `createSpreadsheetCore()` | Cell storage, undo/redo, sort, formula dispatch |
| `SpreadsheetEngine.dll` | `createSpreadsheetModel()` | Virtual `QAbstractTableModel` for QTableView |
| `FileLoader.dll` | `createFileLoader()` | Stream CSV/XLSX in chunks; save back |
| `RibbonUI.dll` | `createRibbonWidget()` | Full ribbon widget, emits formatted signals |

`main.exe` only contains `main()`, `MainWindow`, and `SpreadsheetView`.
All business logic lives inside the DLLs.

---

## Virtual grid — how it works

```
User scrolls ──► QTableView asks model for visible rows
                       │
                       ▼
          SpreadsheetTableModel::data(index, role)
                       │
          Only queries ISpreadsheetCore for cells
          that are VISIBLE in the current viewport
                       │
                       ▼
          SpreadsheetCore returns Cell from
          QHash<(row,col), Cell>  — O(1) lookup
          (never iterates the whole sheet)
```

For loading a 1M-row CSV:
1. `loadMetadata()` counts rows and sets model dimensions instantly.
2. `loadChunk(path, core, sheet, firstRow=0, count=2000)` loads only the top viewport block.
3. On scroll, `MainWindow` (or a future scroll-aware loader) calls `loadChunk` for the next block.
4. Old blocks are evicted from the hash to keep memory bounded.

---

## Formulas

| Category | Functions |
|----------|-----------|
| Math | SUM, AVERAGE/AVG, MIN, MAX, COUNT, COUNTA, ABS, SQRT, POWER, MOD, INT, ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, TRUNC, EXP, LN, LOG10, LOG, PI, SIN, COS, TAN, RAND, RANDBETWEEN |
| Statistical | STDEV, VAR, MEDIAN, LARGE, SMALL |
| Logical | IF, AND, OR, NOT, IFERROR, ISBLANK, ISNUMBER, ISTEXT |
| Text | LEN, LEFT, RIGHT, MID, UPPER, LOWER, TRIM, CONCATENATE, SUBSTITUTE, FIND, VALUE, TEXT, REPT |
| Date | NOW, TODAY, YEAR, MONTH, DAY |
| Lookup | VLOOKUP (basic) |

---

## Generating a large test file

```sh
cd data
python3 generate_sample.py 1000000    # 1 million rows (~120 MB)
python3 generate_sample.py 100000     # 100k rows (~12 MB)
```

Open the resulting `sample_1000k.csv` with **File → Open** and experience the virtual scrolling at scale.

---

## Performance targets

| Metric | Target |
|--------|--------|
| Open 1M-row CSV (metadata + first screen) | < 2 seconds |
| Scroll frame time (1M rows loaded) | < 16 ms |
| Memory for 1M rows (only viewport loaded) | < 50 MB |
| Formula recalc (chain of 1000 dependent cells) | < 100 ms |

---

## Extending with plugins

Create a new DLL that links against `SpreadsheetCore` and implements whatever interface you define:

```cpp
// MyPlugin.dll
extern "C" __declspec(dllexport)
void registerPlugin(ISpreadsheetCore* core) {
    // add custom functions, toolbar actions, etc.
}
```

Load it at runtime with `QLibrary`:
```cpp
QLibrary lib("MyPlugin");
auto fn = (void(*)(ISpreadsheetCore*))lib.resolve("registerPlugin");
if (fn) fn(m_core);
```

---

## License

MIT — see `LICENSE` for details.
