#include "SpreadsheetTableModel.h"
#include "ISpreadsheetCore.h"
#include <QColor>
#include <QFont>
#include <QBrush>
#include <QRegularExpression>
#include <QLocale>

// ── Impl ──────────────────────────────────────────────────────────────────────
struct SpreadsheetTableModel::Impl {
    ISpreadsheetCore* core;
    SheetId           sheet;
    // Minimum visible rows/cols so the grid always shows empty space beyond data
    static constexpr int MIN_ROWS = 200;
    static constexpr int MIN_COLS = 52;  // A–AZ
};

// ── Construction ──────────────────────────────────────────────────────────────
SpreadsheetTableModel::SpreadsheetTableModel(ISpreadsheetCore* core,
                                             SheetId sheet,
                                             QObject* parent)
    : QAbstractTableModel(parent), d(new Impl{core, sheet})
{
    // Incremental cell change — only repaint the changed cell, not the whole sheet
    connect(core, &ISpreadsheetCore::cellChanged,
            this, &SpreadsheetTableModel::onCellChanged);

    // Sheet-level change (sort, insert row etc.) — reset full model
    connect(core, &ISpreadsheetCore::sheetChanged, this, [this](SheetId s) {
        if (s == d->sheet) { beginResetModel(); endResetModel(); }
    });
}

SpreadsheetTableModel::~SpreadsheetTableModel() { delete d; }

// ── Sheet switch ──────────────────────────────────────────────────────────────
void SpreadsheetTableModel::setSheet(SheetId sheet) {
    beginResetModel();
    d->sheet = sheet;
    endResetModel();
}

SheetId SpreadsheetTableModel::currentSheet() const { return d->sheet; }

// ── Dimensions ────────────────────────────────────────────────────────────────
// Always expose at least MIN_ROWS / MIN_COLS so the grid has empty space
// beyond the data (like Excel). QTableView only asks data() for VISIBLE cells,
// so this is zero-cost for empty rows.
int SpreadsheetTableModel::rowCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return qMax(d->MIN_ROWS, d->core->rowCount(d->sheet) + 50);
}

int SpreadsheetTableModel::columnCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return qMax(d->MIN_COLS, d->core->columnCount(d->sheet) + 10);
}

// ── Format number value ───────────────────────────────────────────────────────
static QString formatValue(const QVariant& v, const CellFormat& fmt) {
    if (!v.isValid() || v.isNull()) return {};
    bool ok;
    double d = v.toDouble(&ok);
    if (!ok) return v.toString();

    QLocale locale;
    switch (fmt.numberFormat) {
    case 1: // Number
        return locale.toString(d, 'f', fmt.decimals);
    case 2: // Currency
        return locale.toCurrencyString(d, "$", fmt.decimals);
    case 3: // Percentage
        return locale.toString(d * 100.0, 'f', fmt.decimals) + "%";
    case 4: // Scientific
        return locale.toString(d, 'e', fmt.decimals);
    default: // General
        if (d == (long long)d && qAbs(d) < 1e15)
            return locale.toString((long long)d);
        return locale.toString(d, 'g', 10);
    }
}

// ── Data — only called for VISIBLE cells (virtual / lazy) ────────────────────
QVariant SpreadsheetTableModel::data(const QModelIndex& index, int role) const {
    if (!index.isValid()) return {};
    int r = index.row(), c = index.column();

    // For cells outside the data range, return empty (fast path)
    bool hasData = (r < d->core->rowCount(d->sheet) &&
                    c < d->core->columnCount(d->sheet));
    Cell cell;
    if (hasData) cell = d->core->getCell(d->sheet, r, c);

    switch (role) {
    case Qt::DisplayRole: {
        if (!hasData || (!cell.rawValue.isValid() && !cell.cachedValue.isValid()))
            return {};
        QVariant displayVal = cell.cachedValue.isValid() ? cell.cachedValue : cell.rawValue;
        return formatValue(displayVal, cell.format);
    }
    case Qt::EditRole:
        if (!hasData) return {};
        return cell.formula.isEmpty() ? cell.rawValue : cell.formula;

    case Qt::ForegroundRole:
        if (!hasData) return {};
        return QBrush(cell.format.textColor.isValid()
                      ? cell.format.textColor : QColor("#000000"));

    case Qt::BackgroundRole:
        if (!hasData) return {};
        if (cell.format.fillColor.isValid() && cell.format.fillColor != Qt::white)
            return QBrush(cell.format.fillColor);
        return {};

    case Qt::FontRole: {
        if (!hasData) return {};
        QFont f = cell.format.font;
        if (f.family().isEmpty()) f.setFamily("Calibri");
        if (f.pointSize() <= 0)   f.setPointSize(11);
        f.setBold(cell.format.bold);
        f.setItalic(cell.format.italic);
        f.setUnderline(cell.format.underline);
        return f;
    }
    case Qt::TextAlignmentRole:
        if (!hasData) return (int)(Qt::AlignLeft | Qt::AlignVCenter);
        return (int)cell.format.alignment;

    case Qt::ToolTipRole:
        if (!hasData || cell.formula.isEmpty()) return {};
        return cell.formula;  // Show formula in tooltip

    default:
        return {};
    }
}

// ── setData — called by QTableView when user edits a cell ────────────────────
bool SpreadsheetTableModel::setData(const QModelIndex& index,
                                    const QVariant& value, int role)
{
    if (!index.isValid() || role != Qt::EditRole) return false;
    int r = index.row(), c = index.column();
    QString text = value.toString();
    if (text.startsWith('='))
        d->core->setCellFormula(d->sheet, r, c, text);
    else
        d->core->setCellValue(d->sheet, r, c, value);
    // dataChanged is emitted via onCellChanged signal from core
    return true;
}

// ── Header data ───────────────────────────────────────────────────────────────
QVariant SpreadsheetTableModel::headerData(int section,
                                            Qt::Orientation orientation,
                                            int role) const
{
    if (role == Qt::DisplayRole) {
        if (orientation == Qt::Horizontal)
            return columnLabel(section);   // A, B, ..., Z, AA, AB, ...
        else
            return section + 1;            // 1, 2, 3, ...
    }
    if (role == Qt::TextAlignmentRole)
        return (int)(Qt::AlignCenter);
    return {};
}

// ── Flags ─────────────────────────────────────────────────────────────────────
Qt::ItemFlags SpreadsheetTableModel::flags(const QModelIndex& index) const {
    if (!index.isValid()) return Qt::NoItemFlags;
    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}

// ── Cell changed (incremental update) ────────────────────────────────────────
void SpreadsheetTableModel::onCellChanged(SheetId sheet, int row, int col) {
    if (sheet != d->sheet) return;
    auto idx = index(row, col);
    emit dataChanged(idx, idx, {Qt::DisplayRole, Qt::EditRole,
                                Qt::ForegroundRole, Qt::BackgroundRole,
                                Qt::FontRole, Qt::TextAlignmentRole});
    // If dimensions grew, notify views
    int newRows = qMax(d->MIN_ROWS, d->core->rowCount(d->sheet) + 50);
    int newCols = qMax(d->MIN_COLS, d->core->columnCount(d->sheet) + 10);
    // We don't track old size here; a full reset would be safe but costly.
    // Instead just emit dataChanged for the affected cell — the view will
    // request rowCount/columnCount again on the next layout pass.
    Q_UNUSED(newRows); Q_UNUSED(newCols);
}

// ── Column label utilities ────────────────────────────────────────────────────
QString SpreadsheetTableModel::columnLabel(int col) {
    QString label;
    col++;  // 1-based
    while (col > 0) {
        col--;
        label.prepend(QChar('A' + (col % 26)));
        col /= 26;
    }
    return label;
}

int SpreadsheetTableModel::columnIndex(const QString& label) {
    int col = 0;
    for (QChar ch : label.toUpper())
        col = col * 26 + (ch.unicode() - 'A' + 1);
    return col - 1;
}

// ── Factory ───────────────────────────────────────────────────────────────────
extern "C" ENGINE_API SpreadsheetTableModel*
    createSpreadsheetModel(ISpreadsheetCore* core, SheetId sheet) {
    return new SpreadsheetTableModel(core, sheet);
}



