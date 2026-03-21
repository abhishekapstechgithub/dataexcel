#pragma once
#ifdef _WIN32
  #ifdef SPREADSHEETCORE_EXPORTS
    #define CORE_API __declspec(dllexport)
  #else
    #define CORE_API __declspec(dllimport)
  #endif
#else
  #define CORE_API __attribute__((visibility("default")))
#endif

#include <QObject>
#include <QVariant>
#include <QString>
#include <QColor>
#include <QFont>
#include <memory>

// ── Cell formatting ──────────────────────────────────────────────────────────
struct CORE_API CellFormat {
    QFont   font;
    QColor  textColor   { Qt::black };
    QColor  fillColor   { Qt::white };
    Qt::Alignment alignment { Qt::AlignLeft | Qt::AlignVCenter };
    bool    bold        { false };
    bool    italic      { false };
    bool    underline   { false };
    bool    wrapText    { false };
    int     numberFormat{ 0 };   // 0=General,1=Number,2=Currency,3=Percent,4=Sci
    int     decimals    { 2 };
};

// ── Single cell ──────────────────────────────────────────────────────────────
struct CORE_API Cell {
    QVariant    rawValue;          // what the user typed
    QVariant    cachedValue;       // computed / displayed value
    QString     formula;           // empty if not a formula
    CellFormat  format;
    bool        dirty { true };    // needs recalculation
};

// ── Column metadata ──────────────────────────────────────────────────────────
struct CORE_API ColumnMeta {
    int     width    { 100 };      // pixels
    bool    hidden   { false };
    QString header;                // custom header label (empty = A,B,C,...)
};

// ── Row metadata ─────────────────────────────────────────────────────────────
struct CORE_API RowMeta {
    int   height { 24 };
    bool  hidden { false };
};

// ── Sheet identifier ─────────────────────────────────────────────────────────
using SheetId = int;

class CORE_API ISpreadsheetCore : public QObject
{
    Q_OBJECT
public:
    virtual ~ISpreadsheetCore() = default;

    // ── Cell access ──────────────────────────────────────────────────────────
    virtual Cell    getCell(SheetId sheet, int row, int col) const = 0;
    virtual void    setCell(SheetId sheet, int row, int col, const Cell& cell) = 0;
    virtual void    setCellValue(SheetId sheet, int row, int col, const QVariant& value) = 0;
    virtual void    setCellFormula(SheetId sheet, int row, int col, const QString& formula) = 0;
    virtual void    setCellFormat(SheetId sheet, int row, int col, const CellFormat& fmt) = 0;
    virtual void    clearCell(SheetId sheet, int row, int col) = 0;

    // ── Selection / merge ────────────────────────────────────────────────────
    virtual void    mergeCells(SheetId sheet, int r1, int c1, int r2, int c2) = 0;
    virtual void    unmergeCells(SheetId sheet, int r1, int c1) = 0;

    // ── Row / column ops ─────────────────────────────────────────────────────
    virtual void    insertRow(SheetId sheet, int beforeRow) = 0;
    virtual void    deleteRow(SheetId sheet, int row) = 0;
    virtual void    insertColumn(SheetId sheet, int beforeCol) = 0;
    virtual void    deleteColumn(SheetId sheet, int col) = 0;
    virtual RowMeta    getRowMeta(SheetId sheet, int row) const = 0;
    virtual ColumnMeta getColMeta(SheetId sheet, int col) const = 0;
    virtual void    setRowMeta(SheetId sheet, int row, const RowMeta& m) = 0;
    virtual void    setColMeta(SheetId sheet, int col, const ColumnMeta& m) = 0;

    // ── Dimension ────────────────────────────────────────────────────────────
    virtual int     rowCount(SheetId sheet)    const = 0;
    virtual int     columnCount(SheetId sheet) const = 0;

    // ── Sheet management ─────────────────────────────────────────────────────
    virtual SheetId addSheet(const QString& name = {}) = 0;
    virtual void    removeSheet(SheetId sheet) = 0;
    virtual void    renameSheet(SheetId sheet, const QString& name) = 0;
    virtual QString sheetName(SheetId sheet) const = 0;
    virtual QList<SheetId> sheets() const = 0;

    // ── Undo / redo ──────────────────────────────────────────────────────────
    virtual void    undo() = 0;
    virtual void    redo() = 0;
    virtual bool    canUndo() const = 0;
    virtual bool    canRedo() const = 0;

    // ── Sort / filter ────────────────────────────────────────────────────────
    virtual void    sortRange(SheetId sheet, int r1, int c1, int r2, int c2,
                              int keyCol, Qt::SortOrder order) = 0;

signals:
    void cellChanged(SheetId sheet, int row, int col);
    void sheetChanged(SheetId sheet);
    void formulaError(SheetId sheet, int row, int col, const QString& error);
};

// Factory function exported from the DLL
extern "C" CORE_API ISpreadsheetCore* createSpreadsheetCore();
