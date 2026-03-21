// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetCore.cpp — ISpreadsheetCore adapter over SpreadsheetEngine
//
//  This file bridges the existing ISpreadsheetCore interface (used by the UI)
//  with the new SpreadsheetEngine (which uses the upgraded CellAddress-based
//  sparse storage and dependency-tracking formula evaluator).
//
//  Architecture:
//    UI Layer  →  ISpreadsheetCore (interface)
//                        ↓
//               SpreadsheetCoreImpl (adapter) ← this file
//                        ↓
//               SpreadsheetEngine (engine)    ← engine/src/SpreadsheetEngine.cpp
// ═══════════════════════════════════════════════════════════════════════════════
#include "ISpreadsheetCore.h"
#include "UndoStack.h"
#include "FormulaEngine.h"
#include "../../engine/include/engine/SpreadsheetEngine.h"
#include "../../engine/include/engine/CellAddress.h"
#include <QHash>
#include <QPair>
#include <QMutex>
#include <QSet>
#include <QRegularExpression>
#include <algorithm>

// ── Helper: convert EngineCell → Cell (ISpreadsheetCore Cell) ─────────────────
static Cell toCoreCell(const EngineCell& ec) {
    Cell c;
    c.rawValue    = ec.rawValue;
    c.cachedValue = ec.displayValue.isValid() ? ec.displayValue : ec.rawValue;
    c.formula     = ec.formula;
    c.dirty       = ec.dirty;
    // Convert CellFormat → CellFormat (same struct name, slightly different)
    c.format.bold        = ec.format.bold;
    c.format.italic      = ec.format.italic;
    c.format.underline   = ec.format.underline;
    c.format.wrapText    = ec.format.wrapText;
    c.format.textColor   = ec.format.textColor;
    c.format.fillColor   = ec.format.fillColor;
    c.format.numberFormat= ec.format.numberFmt;
    c.format.decimals    = ec.format.decimals;
    c.format.alignment   =
        (ec.format.hAlign==2 ? Qt::AlignRight :
         ec.format.hAlign==1 ? Qt::AlignHCenter : Qt::AlignLeft) |
        (ec.format.vAlign==0 ? Qt::AlignTop :
         ec.format.vAlign==2 ? Qt::AlignBottom : Qt::AlignVCenter);
    if (!ec.format.fontFamily.isEmpty())
        c.format.font.setFamily(ec.format.fontFamily);
    if (ec.format.fontSize > 0)
        c.format.font.setPointSize(ec.format.fontSize);
    return c;
}

// ── Helper: convert CellFormat (core) → CellFormat (engine) ──────────────────
static EngineCellFormat toEngineFormat(const ::CellFormat& cf) {
    EngineCellFormat ef;
    ef.bold       = cf.bold;
    ef.italic     = cf.italic;
    ef.underline  = cf.underline;
    ef.wrapText   = cf.wrapText;
    ef.textColor  = cf.textColor;
    ef.fillColor  = cf.fillColor;
    ef.numberFmt  = cf.numberFormat;
    ef.decimals   = cf.decimals;
    ef.fontFamily = cf.font.family();
    ef.fontSize   = cf.font.pointSize() > 0 ? cf.font.pointSize() : 11;
    ef.hAlign = (cf.alignment & Qt::AlignRight) ? 2 :
                (cf.alignment & Qt::AlignHCenter) ? 1 : 0;
    ef.vAlign = (cf.alignment & Qt::AlignTop) ? 0 :
                (cf.alignment & Qt::AlignBottom) ? 2 : 1;
    return ef;
}

// ═══════════════════════════════════════════════════════════════════════════════
class SpreadsheetCoreImpl : public ISpreadsheetCore {
    Q_OBJECT
public:
    SpreadsheetCoreImpl() {
        // SpreadsheetEngine already creates "Sheet1" in its constructor
        connect(&m_engine, &SpreadsheetEngine::cellChanged, this,
            [this](EngineSheetId sid, int r, int c){ emit cellChanged((SheetId)sid,r,c); });
        connect(&m_engine, &SpreadsheetEngine::sheetChanged, this,
            [this](EngineSheetId sid){ emit sheetChanged((SheetId)sid); });
        connect(&m_engine, &SpreadsheetEngine::formulaError, this,
            [this](EngineSheetId sid, int r, int c, const QString& e){
                emit formulaError((SheetId)sid,r,c,e); });
    }

    // ── Cell access ──────────────────────────────────────────────────────────
    Cell getCell(SheetId s, int r, int c) const override {
        return toCoreCell(m_engine.getCell((EngineSheetId)s, r, c));
    }
    void setCell(SheetId s, int r, int c, const Cell& cell) override {
        if (!cell.formula.isEmpty())
            m_engine.setCellFormula((EngineSheetId)s, r, c, cell.formula);
        else
            m_engine.setCellValue((EngineSheetId)s, r, c, cell.rawValue);
        m_engine.setCellFormat((EngineSheetId)s, r, c, toEngineFormat(cell.format));
    }
    void setCellValue(SheetId s, int r, int c, const QVariant& v) override {
        m_engine.setCellValue((EngineSheetId)s, r, c, v);
    }
    void setCellFormula(SheetId s, int r, int c, const QString& f) override {
        m_engine.setCellFormula((EngineSheetId)s, r, c, f);
    }
    void setCellFormat(SheetId s, int r, int c, const ::CellFormat& fmt) override {
        m_engine.setCellFormat((EngineSheetId)s, r, c, toEngineFormat(fmt));
    }
    void clearCell(SheetId s, int r, int c) override {
        m_engine.clearCell((EngineSheetId)s, r, c);
    }

    // ── Merge ────────────────────────────────────────────────────────────────
    void mergeCells(SheetId s, int r1, int c1, int r2, int c2) override {
        m_engine.mergeCells((EngineSheetId)s, r1, c1, r2, c2);
    }
    void unmergeCells(SheetId s, int r1, int c1) override {
        m_engine.unmergeCells((EngineSheetId)s, r1, c1);
    }

    // ── Row/col ops ──────────────────────────────────────────────────────────
    void insertRow(SheetId s, int before) override {
        m_engine.insertRows((EngineSheetId)s, before, 1);
    }
    void deleteRow(SheetId s, int row) override {
        m_engine.deleteRows((EngineSheetId)s, row, 1);
    }
    void insertColumn(SheetId s, int before) override {
        m_engine.insertCols((EngineSheetId)s, before, 1);
    }
    void deleteColumn(SheetId s, int col) override {
        m_engine.deleteCols((EngineSheetId)s, col, 1);
    }
    RowMeta    getRowMeta(SheetId s, int r) const override {
        return { m_engine.rowHeight((EngineSheetId)s,r), m_engine.rowHidden((EngineSheetId)s,r) };
    }
    ColumnMeta getColMeta(SheetId s, int c) const override {
        return { m_engine.colWidth((EngineSheetId)s,c), m_engine.colHidden((EngineSheetId)s,c) };
    }
    void setRowMeta(SheetId s, int r, const RowMeta& m) override {
        m_engine.setRowHeight((EngineSheetId)s,r,m.height);
    }
    void setColMeta(SheetId s, int c, const ColumnMeta& m) override {
        m_engine.setColWidth((EngineSheetId)s,c,m.width);
    }

    // ── Dimensions ───────────────────────────────────────────────────────────
    int rowCount(SheetId s)    const override {
        return qMax(200, m_engine.usedRowCount((EngineSheetId)s) + 50);
    }
    int columnCount(SheetId s) const override {
        return qMax(52,  m_engine.usedColCount((EngineSheetId)s) + 10);
    }

    // ── Sheets ───────────────────────────────────────────────────────────────
    SheetId addSheet(const QString& name) override {
        return (SheetId)m_engine.addSheet(name);
    }
    void removeSheet(SheetId s) override { m_engine.removeSheet((EngineSheetId)s); }
    void renameSheet(SheetId s, const QString& n) override { m_engine.renameSheet((EngineSheetId)s,n); }
    QString sheetName(SheetId s) const override { return m_engine.sheetName((EngineSheetId)s); }
    QList<SheetId> sheets() const override {
        QList<SheetId> out;
        for (auto sid : m_engine.sheets()) out.append((SheetId)sid);
        return out;
    }

    // ── Undo/redo ────────────────────────────────────────────────────────────
    void undo() override { m_engine.undo(); }
    void redo() override { m_engine.redo(); }
    bool canUndo() const override { return m_engine.canUndo(); }
    bool canRedo() const override { return m_engine.canRedo(); }

    // ── Sort ─────────────────────────────────────────────────────────────────
    void sortRange(SheetId s, int r1, int c1, int r2, int c2,
                   int keyCol, Qt::SortOrder order) override {
        m_engine.sortRange((EngineSheetId)s,r1,c1,r2,c2,keyCol,
                           order==Qt::AscendingOrder,false);
    }

private:
    SpreadsheetEngine m_engine;
};

#include "SpreadsheetCore.moc"

extern "C" CORE_API ISpreadsheetCore* createSpreadsheetCore() {
    return new SpreadsheetCoreImpl();
}
