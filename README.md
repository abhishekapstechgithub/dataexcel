# OpenSheet

A high-performance desktop spreadsheet application built with C++20 and Qt 6, capable of opening and working with Excel files up to 100GB in size.

## Features

### UI
- **Full ribbon toolbar** with tabs: Home, Insert, Page Layout, Formulas, Data, Review, View, Tools, Premium
- **Formula bar** with cell address box (A1 notation) and fx button
- **Virtual scrolling grid** — rows A to XFD (16,384 columns), 1,048,576 rows displayed
- **Sheet tabs** at the bottom with add (+) button
- **Status bar** with zoom slider (50%–400%) and cell sum display
- Full cell formatting: font, size, bold, italic, underline, text/fill colors, borders, alignment, wrap text, merge cells, number format

### Performance (100GB File Support)
- **Virtual/lazy rendering** — only cells visible on screen are painted; the full file never loads into memory
- **Tile-based cache** — keeps only 10,000 rows × 500 columns in RAM; older tiles are evicted to SQLite
- **Background thread streaming** — UI stays responsive while loading large files; progress bar shown
- **Virtual scrollbar** — maps full 1,048,576-row range (or even 1 billion rows) to the scrollbar
- **SAX-based XLSX parser** — uses `QXmlStreamReader` for forward-only, event-driven parsing (no DOM)

### File Formats
- XLSX read/write (streaming SAX parser — never loads full file into DOM)
- CSV read/write with auto-delimiter detection
- Native `.opensheet` format (ZIP + JSON chunks)

### Formulas (30+)
`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `IF`, `IFERROR`, `AND`, `OR`, `NOT`, `IFS`,
`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT`,
`COUNTIF`, `SUMIF`, `AVERAGEIF`, `STDEV`, `VAR`, `MEDIAN`,
`CONCATENATE`, `LEN`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `UPPER`, `LOWER`,
`TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, and more.

### Other Features
- Find & Replace with regex support
- Column/row resize by dragging headers
- Freeze panes (row and/or column)
- Multiple sheets (add, rename, delete, reorder)
- Undo/Redo (100+ levels)
- Copy/paste with system clipboard (tab-delimited text)
- Sort ascending/descending
- Charts: Bar, Line, Pie, Scatter (via Qt6::Charts)
- Drag & drop file opening

## Build

### Prerequisites
- CMake 3.20+
- Qt 6.6+ (Widgets, Core, Gui, Concurrent, Xml, Charts, Sql)
- MSVC 2019+ or Clang/GCC with C++20 support

### Windows (MSVC)
```bat
build.bat
```

Or manually:
```bat
cmake -B build -DCMAKE_PREFIX_PATH="C:/Qt/6.6.0/msvc2019_64"
cmake --build build --config Release
cd build/bin
windeployqt OpenSheet.exe
```

### Linux/macOS
```sh
cmake -B build -DCMAKE_BUILD_TYPE=Release -DCMAKE_PREFIX_PATH=/opt/Qt/6.6/gcc_64
cmake --build build --parallel
```

## CI/CD

GitHub Actions automatically builds a Windows x64 installer on every push to `main`.
The installer is built with **Inno Setup 6** and available as a GitHub Actions artifact.

Download the latest build from the **Actions** tab → **OpenSheet Windows x64 Installer**.

## Architecture

```
OpenSheet/
├── app/              # Main executable
│   ├── main.cpp
│   ├── MainWindow.cpp/h         # Main window (Excel-like layout)
│   ├── SpreadsheetView.cpp/h    # Virtual grid renderer
│   ├── FormulaBar.cpp/h         # Address box + formula editor
│   ├── SheetTabBar.cpp/h        # Sheet tabs
│   ├── TileCache.cpp/h          # 10K×500 tile cache (SQLite-backed)
│   └── FindReplaceDialog/FormatCellsDialog
├── dll/
│   ├── FormulaEngine/           # Tokenizer + Parser + 40+ functions
│   ├── FileLoader/              # CSV + XLSX SAX streaming loader
│   ├── SpreadsheetCore/         # ISpreadsheetCore adapter + UndoStack
│   ├── SpreadsheetEngine/       # QAbstractTableModel bridge
│   └── RibbonUI/                # 9-tab Excel-style ribbon
├── engine/                      # Core: sparse cell storage + formula dependency graph
├── installer/
│   └── OpenSheet.iss            # Inno Setup 6 script
└── .github/workflows/
    └── build-windows.yml        # GitHub Actions CI
```

## License

MIT — see LICENSE file.
