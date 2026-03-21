#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  EngineModel.h — QAbstractTableModel backed by SpreadsheetEngine
//  - Presents a virtual grid of MAX_ROWS×MAX_COLS to QTableView
//  - Only asks the engine for data() of VISIBLE cells (viewport-based lazy load)
//  - Responds to engine signals for targeted cell repaints (not full reset)
// ═══════════════════════════════════════════════════════════════════════════════
#include "SpreadsheetEngine.h"
#include <QAbstractTableModel>
#include <QLocale>
#include <QDate>
#include <QDateTime>

class EngineModel : public QAbstractTableModel {
    Q_OBJECT
public:
    explicit EngineModel(SpreadsheetEngine* engine,
                         EngineSheetId sheet,
                         QObject* parent = nullptr);

    void setSheet(EngineSheetId sheet);
    EngineSheetId currentSheet() const { return m_sheet; }

    // QAbstractTableModel overrides
    int      rowCount   (const QModelIndex& = {}) const override;
    int      columnCount(const QModelIndex& = {}) const override;
    QVariant data       (const QModelIndex& idx, int role = Qt::DisplayRole) const override;
    bool     setData    (const QModelIndex& idx, const QVariant& value, int role = Qt::EditRole) override;
    QVariant headerData (int section, Qt::Orientation, int role = Qt::DisplayRole) const override;
    Qt::ItemFlags flags (const QModelIndex& idx) const override;

    // Column label utilities
    static QString  columnLabel(int col);
    static int      columnIndex(const QString& label);

    // Formatted display value
    static QString  formatValue(const QVariant& v, const EngineCellFormat& fmt);

private slots:
    void onCellChanged(EngineSheetId sheet, int row, int col);
    void onSheetChanged(EngineSheetId sheet);

private:
    SpreadsheetEngine* m_engine;
    EngineSheetId      m_sheet;

    // Grid always shows at least these many rows/cols
    static constexpr int VIRTUAL_ROWS = 1'048'576;
    static constexpr int VIRTUAL_COLS = 16'384;
    static constexpr int MIN_DISPLAY_ROWS = 200;
    static constexpr int MIN_DISPLAY_COLS = 52;
};
