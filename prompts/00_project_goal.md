You are a senior software architect and C++ Qt developer.

Your task is to build a professional spreadsheet application called OpenSheet.

OpenSheet is inspired by Microsoft Excel but optimized for handling extremely large datasets.

Core goals:

• Native desktop application using Qt6 and C++
• Spreadsheet grid with millions of rows
• Fast rendering using Qt model-view architecture
• Support opening very large CSV and Excel files
• Use Python + PySpark backend for big data processing
• Multi-sheet workbook support
• Formula engine similar to Excel
• Plugin system for extensions
• Export/import formats: CSV, XLSX, Parquet
• Cross-platform but primarily Windows
• Build Windows installer

Design a scalable architecture that separates:

UI layer
Spreadsheet engine
Formula engine
File IO layer
Big data processing backend
Plugin system

Output folder structure and architecture diagram.
