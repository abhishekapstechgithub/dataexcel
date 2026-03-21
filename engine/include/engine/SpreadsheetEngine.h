#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetEngine.h — Core data engine
//  Responsibilities:
//    - Sparse cell storage (unordered_map) for 1M×16K grid
//    - Formula evaluation with dependency graph
//    - Auto-recalculation when cells change
//    - Sheet management
//    - Undo/redo
//  Thread safety: public API is main-thread only; recalc can run on worker thread
// ═══════════════════════════════════════════════════════════════════════════════
#include "CellAddress.h"
#include <QObject>
#include <QVariant>
#include <QString>
#include <QColor>
#include <QFont>
#include <QList>
#include <QHash>
#include <unordered_map>
#include <functional>
#include <memory>

// ── Cell data types ───────────────────────────────────────────────────────────
enum class CellType { Empty, Text, Number, Boolean, Formula, Error };

struct EngineCellFormat {
    QString fontFamily   { "Calibri" };
    int     fontSize     { 11 };
    bool    bold         { false };
    bool    italic       { false };
    bool    underline    { false };
    QColor  textColor    { Qt::black };
    QColor  fillColor    { Qt::white };
    int     hAlign       { 0 };   // 0=Left 1=Center 2=Right
    int     vAlign       { 1 };   // 0=Top  1=Mid    2=Bottom
    bool    wrapText     { false };
    int     numberFmt    { 0 };   // 0=General 1=Number 2=Currency 3=Percent 4=Sci
    int     decimals     { 2 };
    bool    strikethrough{ false };

    bool operator==(const EngineCellFormat& o) const = default;
};

struct EngineCell {
    CellType  type        { CellType::Empty };
    QVariant  rawValue;           // what the user typed / stored
    QVariant  displayValue;       // computed/formatted for display
    QString   formula;            // non-empty if this is a formula cell
    EngineCellFormat format;
    bool      dirty       { false }; // needs recalculation

    bool isEmpty() const { return type == CellType::Empty; }
};

// ── Sheet identifier ──────────────────────────────────────────────────────────
using EngineSheetId = int;

// ── Engine interface ──────────────────────────────────────────────────────────
class SpreadsheetEngine : public QObject {
    Q_OBJECT
public:
    explicit SpreadsheetEngine(QObject* parent = nullptr);
    ~SpreadsheetEngine() override;

    // ── Cell access ───────────────────────────────────────────────────────────
    EngineCell  getCell(EngineSheetId sheet, int row, int col) const;
    void        setCellValue(EngineSheetId sheet, int row, int col,
                             const QVariant& value);
    void        setCellFormula(EngineSheetId sheet, int row, int col,
                               const QString& formula);
    void        setCellFormat(EngineSheetId sheet, int row, int col,
                              const EngineCellFormat& fmt);
    void        clearCell(EngineSheetId sheet, int row, int col);
    void        clearRange(EngineSheetId sheet,
                           int r1, int c1, int r2, int c2);

    // ── Bulk operations (used by file loader) ─────────────────────────────────
    // Fast path: bypasses undo stack, used during file load
    void        bulkSetValue(EngineSheetId sheet, int row, int col,
                             const QVariant& value);
    void        bulkFinalize(EngineSheetId sheet);  // recalc all formulas

    // ── Dimensions ────────────────────────────────────────────────────────────
    int  usedRowCount(EngineSheetId sheet) const;    // highest used row + 1
    int  usedColCount(EngineSheetId sheet) const;    // highest used col + 1
    bool cellExists(EngineSheetId sheet, int row, int col) const;

    // ── Row/col metadata ──────────────────────────────────────────────────────
    int  rowHeight(EngineSheetId sheet, int row) const;
    int  colWidth(EngineSheetId sheet, int col) const;
    void setRowHeight(EngineSheetId sheet, int row, int h);
    void setColWidth(EngineSheetId sheet, int col, int w);
    bool rowHidden(EngineSheetId sheet, int row) const;
    bool colHidden(EngineSheetId sheet, int col) const;

    // ── Row / column insert / delete ──────────────────────────────────────────
    void insertRows(EngineSheetId sheet, int before, int count = 1);
    void deleteRows(EngineSheetId sheet, int row,    int count = 1);
    void insertCols(EngineSheetId sheet, int before, int count = 1);
    void deleteCols(EngineSheetId sheet, int col,    int count = 1);

    // ── Merge ─────────────────────────────────────────────────────────────────
    void mergeCells(EngineSheetId sheet, int r1, int c1, int r2, int c2);
    void unmergeCells(EngineSheetId sheet, int r1, int c1);
    bool isMerged(EngineSheetId sheet, int row, int col) const;
    CellAddress mergeOrigin(EngineSheetId sheet, int row, int col) const;

    // ── Sheet management ──────────────────────────────────────────────────────
    EngineSheetId addSheet(const QString& name = {});
    void          removeSheet(EngineSheetId sheet);
    void          renameSheet(EngineSheetId sheet, const QString& name);
    void          moveSheet(EngineSheetId sheet, int newIndex);
    QString       sheetName(EngineSheetId sheet) const;
    QList<EngineSheetId> sheets() const;
    int           sheetIndex(EngineSheetId sheet) const;

    // ── Sort ──────────────────────────────────────────────────────────────────
    void sortRange(EngineSheetId sheet,
                   int r1, int c1, int r2, int c2,
                   int keyCol, bool ascending, bool hasHeader = false);

    // ── Undo / redo ───────────────────────────────────────────────────────────
    void  undo();
    void  redo();
    bool  canUndo() const;
    bool  canRedo() const;

    // ── Formula recalc ────────────────────────────────────────────────────────
    void  recalcAll(EngineSheetId sheet);
    void  recalcCell(EngineSheetId sheet, int row, int col);

    // ── Search ────────────────────────────────────────────────────────────────
    struct SearchResult { EngineSheetId sheet; int row; int col; };
    QList<SearchResult> findAll(EngineSheetId sheet,
                                const QString& text,
                                bool matchCase = false,
                                bool wholeCell = false) const;

signals:
    void cellChanged(EngineSheetId sheet, int row, int col);
    void rangeChanged(EngineSheetId sheet, int r1, int c1, int r2, int c2);
    void sheetChanged(EngineSheetId sheet);
    void formulaError(EngineSheetId sheet, int row, int col, const QString& err);
    void recalcFinished(EngineSheetId sheet);

private:
    struct Impl;
    std::unique_ptr<Impl> d;
};
