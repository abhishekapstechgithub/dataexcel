Build the Qt6 user interface for OpenSheet.

Requirements:

Ribbon interface similar to Excel.

Tabs:

Home
Insert
Formulas
Data
View

Components:

SpreadsheetView (QTableView based)
WorkbookTabs
FormulaBar
StatusBar

Rendering requirements:

• Virtual scrolling
• Efficient cell repainting
• 1 million+ rows support

Use:

QAbstractTableModel
QStyledItemDelegate
Qt Graphics optimization

Generate:

MainWindow
RibbonBar
SpreadsheetView
Workbook class
