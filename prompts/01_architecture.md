Design the full architecture for OpenSheet.

Components:

1. Qt UI Layer
   Ribbon interface
   Spreadsheet grid view
   Workbook manager
   Formula bar

2. Spreadsheet Engine (C++)
   Cell storage
   Row/column virtualization
   Format system
   Dependency graph

3. File IO Engine
   CSV reader
   XLSX reader
   Parquet reader

4. Python Data Engine
   PySpark integration
   Lazy loading large datasets
   Chunk-based loading

5. Formula Engine
   SUM
   AVERAGE
   COUNT
   IF
   VLOOKUP
   pivot table support

6. Memory optimization
   Columnar storage
   lazy evaluation

Provide folder structure and interfaces.
