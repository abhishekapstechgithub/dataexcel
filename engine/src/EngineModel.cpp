// ═══════════════════════════════════════════════════════════════════════════════
//  EngineModel.cpp — Virtual QAbstractTableModel over SpreadsheetEngine
//
//  Performance principles:
//  1. rowCount/columnCount return the virtual max (1M × 16K) so the view can
//     scroll freely, but QTableView only calls data() for visible rows.
//  2. data() is O(1) — hash lookup in the engine's sparse map.
//  3. Cell changes trigger dataChanged() for only the changed cell (no full reset).
//  4. Sheet changes trigger beginResetModel/endResetModel.
// ═══════════════════════════════════════════════════════════════════════════════
#include "../include/engine/EngineModel.h"
#include <QBrush>
#include <QFont>
#include <QLocale>
#include <QDate>

EngineModel::EngineModel(SpreadsheetEngine* engine, EngineSheetId sheet, QObject* parent)
    : QAbstractTableModel(parent), m_engine(engine), m_sheet(sheet)
{
    connect(engine, &SpreadsheetEngine::cellChanged,  this, &EngineModel::onCellChanged);
    connect(engine, &SpreadsheetEngine::sheetChanged, this, &EngineModel::onSheetChanged);
}

void EngineModel::setSheet(EngineSheetId sheet) {
    beginResetModel();
    m_sheet = sheet;
    endResetModel();
}

// ── Dimensions ────────────────────────────────────────────────────────────────
// QTableView queries this once; the virtual size allows scrolling to any cell.
int EngineModel::rowCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return qMax(MIN_DISPLAY_ROWS,
                m_engine->usedRowCount(m_sheet) + 50);
}

int EngineModel::columnCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return qMax(MIN_DISPLAY_COLS,
                m_engine->usedColCount(m_sheet) + 10);
}

// ── data ─────────────────────────────────────────────────────────────────────
// Called ONLY for visible cells — the core of the lazy-loading design.
QVariant EngineModel::data(const QModelIndex& idx, int role) const {
    if (!idx.isValid()) return {};
    int r = idx.row(), c = idx.column();

    // Fast path: if cell doesn't exist in sparse storage, return defaults
    bool exists = m_engine->cellExists(m_sheet, r, c);
    EngineCell cell;
    if (exists) cell = m_engine->getCell(m_sheet, r, c);

    switch (role) {
    case Qt::DisplayRole: {
        if (!exists || cell.isEmpty()) return {};
        QVariant dv = cell.displayValue.isValid() ? cell.displayValue : cell.rawValue;
        return formatValue(dv, cell.format);
    }
    case Qt::EditRole:
        if (!exists) return {};
        return cell.formula.isEmpty() ? cell.rawValue : cell.formula;

    case Qt::ForegroundRole:
        if (!exists) return {};
        return QBrush(cell.format.textColor);

    case Qt::BackgroundRole:
        if (!exists || cell.format.fillColor == Qt::white) return {};
        return QBrush(cell.format.fillColor);

    case Qt::FontRole: {
        if (!exists) return {};
        QFont f;
        f.setFamily(cell.format.fontFamily.isEmpty() ? "Calibri" : cell.format.fontFamily);
        f.setPointSize(cell.format.fontSize > 0 ? cell.format.fontSize : 11);
        f.setBold(cell.format.bold);
        f.setItalic(cell.format.italic);
        f.setUnderline(cell.format.underline);
        f.setStrikeOut(cell.format.strikethrough);
        return f;
    }
    case Qt::TextAlignmentRole: {
        if (!exists) return (int)(Qt::AlignLeft | Qt::AlignVCenter);
        Qt::Alignment h = (cell.format.hAlign == 2) ? Qt::AlignRight :
                          (cell.format.hAlign == 1) ? Qt::AlignHCenter : Qt::AlignLeft;
        Qt::Alignment v = (cell.format.vAlign == 0) ? Qt::AlignTop :
                          (cell.format.vAlign == 2) ? Qt::AlignBottom : Qt::AlignVCenter;
        return (int)(h | v);
    }
    case Qt::ToolTipRole:
        if (!exists || cell.formula.isEmpty()) return {};
        return cell.formula;

    default:
        return {};
    }
}

// ── setData ───────────────────────────────────────────────────────────────────
bool EngineModel::setData(const QModelIndex& idx, const QVariant& value, int role) {
    if (!idx.isValid() || role != Qt::EditRole) return false;
    QString text = value.toString().trimmed();
    if (text.startsWith('='))
        m_engine->setCellFormula(m_sheet, idx.row(), idx.column(), text);
    else
        m_engine->setCellValue(m_sheet, idx.row(), idx.column(), value);
    return true;
}

// ── headerData ────────────────────────────────────────────────────────────────
QVariant EngineModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role == Qt::DisplayRole) {
        if (orientation == Qt::Horizontal) return columnLabel(section);
        return section + 1;
    }
    if (role == Qt::TextAlignmentRole) return (int)Qt::AlignCenter;
    return {};
}

Qt::ItemFlags EngineModel::flags(const QModelIndex& idx) const {
    if (!idx.isValid()) return Qt::NoItemFlags;
    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}

// ── Slots ─────────────────────────────────────────────────────────────────────
void EngineModel::onCellChanged(EngineSheetId sheet, int row, int col) {
    if (sheet != m_sheet) return;
    auto idx = index(row, col);
    emit dataChanged(idx, idx, {Qt::DisplayRole, Qt::EditRole,
                                Qt::ForegroundRole, Qt::BackgroundRole,
                                Qt::FontRole, Qt::TextAlignmentRole});
}

void EngineModel::onSheetChanged(EngineSheetId sheet) {
    if (sheet != m_sheet) return;
    beginResetModel(); endResetModel();
}

// ── Utilities ─────────────────────────────────────────────────────────────────
QString EngineModel::columnLabel(int col) {
    QString label; ++col;
    while (col > 0) { --col; label.prepend(QChar('A'+col%26)); col/=26; }
    return label;
}

int EngineModel::columnIndex(const QString& label) {
    int col=0;
    for (QChar ch : label.toUpper()) col = col*26 + (ch.unicode()-'A'+1);
    return col-1;
}

QString EngineModel::formatValue(const QVariant& v, const EngineCellFormat& fmt) {
    if (!v.isValid() || v.isNull()) return {};
    bool ok;
    double d = v.toDouble(&ok);
    if (!ok) return v.toString();

    QLocale locale;
    switch (fmt.numberFmt) {
    case 1: return locale.toString(d,'f',fmt.decimals);
    case 2: return locale.toCurrencyString(d,"$",fmt.decimals);
    case 3: return locale.toString(d*100.0,'f',fmt.decimals)+"%";
    case 4: return locale.toString(d,'e',fmt.decimals);
    case 5: // Date — avoid QDate include dependency; just format as string
        return v.toString();
    default: // General
        if (d==(long long)d && qAbs(d)<1e15)
            return locale.toString((long long)d);
        return locale.toString(d,'g',10);
    }
}
