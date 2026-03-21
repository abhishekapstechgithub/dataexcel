#include "SpreadsheetTableModel.h"
#include "ISpreadsheetCore.h"
#include <QColor>
#include <QFont>
#include <QBrush>
#include <QRegularExpression>

// ── Impl ──────────────────────────────────────────────────────────────────────
struct SpreadsheetTableModel::Impl {
    ISpreadsheetCore* core;
    SheetId           sheet;
};

// ── Construction ──────────────────────────────────────────────────────────────
SpreadsheetTableModel::SpreadsheetTableModel(ISpreadsheetCore* core,
                                             SheetId sheet,
                                             QObject* parent)
    : QAbstractTableModel(parent), d(new Impl{core, sheet})
{
    connect(core, &ISpreadsheetCore::cellChanged,
            this, &SpreadsheetTableModel::onCellChanged);
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
int SpreadsheetTableModel::rowCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return d->core->rowCount(d->sheet);
}

int SpreadsheetTableModel::columnCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return d->core->columnCount(d->sheet);
}

// ── Data — only called for VISIBLE cells by QTableView ────────────────────────
QVariant SpreadsheetTableModel::data(const QModelIndex& index, int role) const {
    if (!index.isValid()) return {};
    int r = index.row(), c = index.column();
    Cell cell = d->core->getCell(d->sheet, r, c);

    switch (role) {
    case Qt::DisplayRole:
    case Qt::EditRole: {
        if (role == Qt::EditRole)
            return cell.formula.isEmpty() ? cell.rawValue : cell.formula;
        // Format displayed value
        QVariant val = cell.cachedValue.isValid() ? cell.cachedValue : cell.rawValue;
        if (!val.isValid() || val.isNull()) return {};
        const CellFormat& fmt = cell.format;
        bool ok;
        double dbl = val.toDouble(&ok);
        if (ok) {
            switch (fmt.numberFormat) {
            case 1: return QString("%L1").arg(dbl, 0, 'f', fmt.decimals);
            case 2: return QString("$%L1").arg(dbl, 0, 'f', fmt.decimals);
            case 3: return QString("%1%").arg(dbl * 100.0, 0, 'f', fmt.decimals);
            case 4: return QString("%1").arg(dbl, 0, 'e', fmt.decimals);
            default: break;
            }
        }
        return val.toString();
    }
    case Qt::FontRole: {
        QFont f = cell.format.font;
        f.setBold(cell.format.bold);
        f.setItalic(cell.format.italic);
        f.setUnderline(cell.format.underline);
        return f;
    }
    case Qt::ForegroundRole:
        return QBrush(cell.format.textColor);
    case Qt::BackgroundRole:
        if (cell.format.fillColor != Qt::white)
            return QBrush(cell.format.fillColor);
        return {};
    case Qt::TextAlignmentRole:
        return QVariant(cell.format.alignment);
    case Qt::ToolTipRole:
        return cell.formula.isEmpty() ? cell.rawValue.toString() : cell.formula;
    default:
        return {};
    }
}

// ── setData ───────────────────────────────────────────────────────────────────
bool SpreadsheetTableModel::setData(const QModelIndex& index,
                                    const QVariant& value, int role)
{
    if (!index.isValid() || role != Qt::EditRole) return false;
    int r = index.row(), c = index.column();
    QString str = value.toString().trimmed();

    if (str.startsWith('='))
        d->core->setCellFormula(d->sheet, r, c, str);
    else
        d->core->setCellValue(d->sheet, r, c, value);

    emit dataChanged(index, index);
    return true;
}

// ── Header ────────────────────────────────────────────────────────────────────
QVariant SpreadsheetTableModel::headerData(int section,
                                           Qt::Orientation orientation,
                                           int role) const
{
    if (role != Qt::DisplayRole) return {};
    if (orientation == Qt::Horizontal) {
        ColumnMeta m = d->core->getColMeta(d->sheet, section);
        return m.header.isEmpty() ? columnLabel(section) : m.header;
    } else {
        return QString::number(section + 1);
    }
}

Qt::ItemFlags SpreadsheetTableModel::flags(const QModelIndex& index) const {
    if (!index.isValid()) return Qt::NoItemFlags;
    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}

// ── Cell changed slot ─────────────────────────────────────────────────────────
void SpreadsheetTableModel::onCellChanged(SheetId sheet, int row, int col) {
    if (sheet != d->sheet) return;
    auto idx = createIndex(row, col);
    emit dataChanged(idx, idx, {Qt::DisplayRole, Qt::EditRole});
}

// ── Static helpers ────────────────────────────────────────────────────────────
QString SpreadsheetTableModel::columnLabel(int col) {
    QString label;
    col++;
    while (col > 0) {
        col--;
        label.prepend(QChar('A' + (col % 26)));
        col /= 26;
    }
    return label;
}

int SpreadsheetTableModel::columnIndex(const QString& label) {
    int idx = 0;
    for (QChar c : label.toUpper())
        idx = idx * 26 + (c.unicode() - 'A' + 1);
    return idx - 1;
}

// ── Factory ───────────────────────────────────────────────────────────────────
extern "C" ENGINE_API SpreadsheetTableModel*
    createSpreadsheetModel(ISpreadsheetCore* core, SheetId sheet) {
    return new SpreadsheetTableModel(core, sheet);
}
