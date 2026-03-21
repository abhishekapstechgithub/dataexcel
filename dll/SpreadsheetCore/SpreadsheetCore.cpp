#include "ISpreadsheetCore.h"
#include "UndoStack.h"
#include "FormulaEngine.h"
#include <QHash>
#include <QPair>
#include <QMutex>
#include <QSet>
#include <QRegularExpression>
#include <algorithm>

// ── Key types ─────────────────────────────────────────────────────────────────
using CellKey = QPair<int,int>; // (row, col)
using SheetData = QHash<CellKey, Cell>;

struct MergeRange { int r1,c1,r2,c2; };

struct SheetInfo {
    QString                  name;
    SheetData                cells;
    QHash<int, RowMeta>      rows;
    QHash<int, ColumnMeta>   cols;
    QList<MergeRange>        merges;
    int                      maxRow { 0 };
    int                      maxCol { 0 };
};

// ── Concrete implementation ────────────────────────────────────────────────────
class SpreadsheetCoreImpl : public ISpreadsheetCore
{
    Q_OBJECT
public:
    SpreadsheetCoreImpl() {
        addSheet("Sheet1");
    }

    // ── Cell access ──────────────────────────────────────────────────────────
    Cell getCell(SheetId s, int r, int c) const override {
        auto it = m_sheets.find(s);
        if (it == m_sheets.end()) return Cell{};
        return it->cells.value({r,c}, Cell{});
    }

    void setCell(SheetId s, int r, int c, const Cell& cell) override {
        ensureSheet(s);
        Cell old = m_sheets[s].cells.value({r,c});
        m_undoStack.push({
            [=, this]() mutable {
                m_sheets[s].cells[{r,c}] = cell;
                updateDims(s, r, c);
                recalcDependents(s, r, c);
                emit cellChanged(s, r, c);
            },
            [=, this]() mutable {
                m_sheets[s].cells[{r,c}] = old;
                recalcDependents(s, r, c);
                emit cellChanged(s, r, c);
            },
            "Set cell"
        });
    }

    void setCellValue(SheetId s, int r, int c, const QVariant& value) override {
        Cell cell = getCell(s,r,c);
        cell.rawValue = value;
        cell.formula.clear();
        cell.cachedValue = value;
        cell.dirty = false;
        setCell(s, r, c, cell);
    }

    void setCellFormula(SheetId s, int r, int c, const QString& formula) override {
        Cell cell = getCell(s,r,c);
        cell.formula = formula;
        cell.rawValue = formula;
        cell.dirty = true;
        recalcCell(s, r, c, cell);
        setCell(s, r, c, cell);
    }

    void setCellFormat(SheetId s, int r, int c, const CellFormat& fmt) override {
        Cell cell = getCell(s,r,c);
        CellFormat old = cell.format;
        m_undoStack.push({
            [=, this]() mutable { m_sheets[s].cells[{r,c}].format = fmt; emit cellChanged(s,r,c); },
            [=, this]() mutable { m_sheets[s].cells[{r,c}].format = old; emit cellChanged(s,r,c); },
            "Format cell"
        });
    }

    void clearCell(SheetId s, int r, int c) override {
        setCellValue(s, r, c, QVariant());
    }

    // ── Merge ────────────────────────────────────────────────────────────────
    void mergeCells(SheetId s, int r1, int c1, int r2, int c2) override {
        ensureSheet(s);
        m_sheets[s].merges.append({r1,c1,r2,c2});
        emit sheetChanged(s);
    }

    void unmergeCells(SheetId s, int r1, int c1) override {
        ensureSheet(s);
        auto& merges = m_sheets[s].merges;
        merges.erase(std::remove_if(merges.begin(), merges.end(),
            [r1,c1](const MergeRange& m){ return m.r1==r1 && m.c1==c1; }), merges.end());
        emit sheetChanged(s);
    }

    // ── Row/col ops ──────────────────────────────────────────────────────────
    void insertRow(SheetId s, int before) override {
        ensureSheet(s);
        SheetData newData;
        for (auto it = m_sheets[s].cells.begin(); it != m_sheets[s].cells.end(); ++it) {
            int r = it.key().first, c = it.key().second;
            newData[{r >= before ? r+1 : r, c}] = it.value();
        }
        m_sheets[s].cells = newData;
        m_sheets[s].maxRow++;
        emit sheetChanged(s);
    }

    void deleteRow(SheetId s, int row) override {
        ensureSheet(s);
        SheetData newData;
        for (auto it = m_sheets[s].cells.begin(); it != m_sheets[s].cells.end(); ++it) {
            int r = it.key().first, c = it.key().second;
            if (r == row) continue;
            newData[{r > row ? r-1 : r, c}] = it.value();
        }
        m_sheets[s].cells = newData;
        if (m_sheets[s].maxRow > 0) m_sheets[s].maxRow--;
        emit sheetChanged(s);
    }

    void insertColumn(SheetId s, int before) override {
        ensureSheet(s);
        SheetData newData;
        for (auto it = m_sheets[s].cells.begin(); it != m_sheets[s].cells.end(); ++it) {
            int r = it.key().first, c = it.key().second;
            newData[{r, c >= before ? c+1 : c}] = it.value();
        }
        m_sheets[s].cells = newData;
        m_sheets[s].maxCol++;
        emit sheetChanged(s);
    }

    void deleteColumn(SheetId s, int col) override {
        ensureSheet(s);
        SheetData newData;
        for (auto it = m_sheets[s].cells.begin(); it != m_sheets[s].cells.end(); ++it) {
            int r = it.key().first, c = it.key().second;
            if (c == col) continue;
            newData[{r, c > col ? c-1 : c}] = it.value();
        }
        m_sheets[s].cells = newData;
        if (m_sheets[s].maxCol > 0) m_sheets[s].maxCol--;
        emit sheetChanged(s);
    }

    RowMeta    getRowMeta(SheetId s, int r)    const override { return m_sheets.value(s).rows.value(r, RowMeta{}); }
    ColumnMeta getColMeta(SheetId s, int c)    const override { return m_sheets.value(s).cols.value(c, ColumnMeta{}); }
    void setRowMeta(SheetId s, int r, const RowMeta& m)    override { ensureSheet(s); m_sheets[s].rows[r] = m; }
    void setColMeta(SheetId s, int c, const ColumnMeta& m) override { ensureSheet(s); m_sheets[s].cols[c] = m; }

    // ── Dimensions ───────────────────────────────────────────────────────────
    int rowCount(SheetId s)    const override { return qMax(m_sheets.value(s).maxRow + 1, 1000000); }
    int columnCount(SheetId s) const override { return qMax(m_sheets.value(s).maxCol + 1, 1024); }

    // ── Sheets ───────────────────────────────────────────────────────────────
    SheetId addSheet(const QString& name) override {
        SheetId id = m_nextSheetId++;
        m_sheets[id].name = name.isEmpty() ? QString("Sheet%1").arg(id+1) : name;
        m_sheetOrder.append(id);
        return id;
    }

    void removeSheet(SheetId s) override {
        m_sheets.remove(s);
        m_sheetOrder.removeAll(s);
    }

    void renameSheet(SheetId s, const QString& name) override {
        if (m_sheets.contains(s)) m_sheets[s].name = name;
    }

    QString sheetName(SheetId s) const override { return m_sheets.value(s).name; }
    QList<SheetId> sheets() const override { return m_sheetOrder; }

    // ── Undo/redo ────────────────────────────────────────────────────────────
    void undo() override { m_undoStack.undo(); }
    void redo() override { m_undoStack.redo(); }
    bool canUndo() const override { return m_undoStack.canUndo(); }
    bool canRedo() const override { return m_undoStack.canRedo(); }

    // ── Sort ─────────────────────────────────────────────────────────────────
    void sortRange(SheetId s, int r1, int c1, int r2, int c2,
                   int keyCol, Qt::SortOrder order) override {
        ensureSheet(s);
        // Collect rows in range
        QList<QHash<int,Cell>> rows;
        for (int r = r1; r <= r2; ++r) {
            QHash<int,Cell> row;
            for (int c = c1; c <= c2; ++c)
                row[c] = m_sheets[s].cells.value({r,c});
            rows.append(row);
        }
        std::stable_sort(rows.begin(), rows.end(),
            [keyCol, order](const QHash<int,Cell>& a, const QHash<int,Cell>& b) {
                QVariant va = a.value(keyCol).cachedValue;
                QVariant vb = b.value(keyCol).cachedValue;
                bool aIsNum, bIsNum;
                double da = va.toDouble(&aIsNum), db = vb.toDouble(&bIsNum);
                bool less = (aIsNum && bIsNum) ? da < db : va.toString() < vb.toString();
                return order == Qt::AscendingOrder ? less : !less;
            });
        for (int r = r1; r <= r2; ++r) {
            const auto& row = rows[r - r1];
            for (int c = c1; c <= c2; ++c)
                m_sheets[s].cells[{r,c}] = row.value(c);
        }
        emit sheetChanged(s);
    }

private:
    QHash<SheetId, SheetInfo> m_sheets;
    QList<SheetId>            m_sheetOrder;
    SheetId                   m_nextSheetId { 0 };
    UndoStack                 m_undoStack;
    FormulaEngine             m_formula;

    void ensureSheet(SheetId s) {
        if (!m_sheets.contains(s)) m_sheets[s] = SheetInfo{};
    }

    void updateDims(SheetId s, int r, int c) {
        m_sheets[s].maxRow = qMax(m_sheets[s].maxRow, r);
        m_sheets[s].maxCol = qMax(m_sheets[s].maxCol, c);
    }

    // Resolve a cell reference string like "A1" to its cached value
    QVariant resolveCell(SheetId s, const QString& ref) {
        static QRegularExpression re(R"(([A-Z]+)(\d+))");
        auto m = re.match(ref);
        if (!m.hasMatch()) return QVariant();
        int col = 0;
        for (QChar ch : m.captured(1)) col = col * 26 + (ch.unicode() - 'A' + 1);
        col--;
        int row = m.captured(2).toInt() - 1;
        Cell c = getCell(s, row, col);
        return c.dirty ? recalcAndGet(s, row, col) : c.cachedValue;
    }

    QVariant recalcAndGet(SheetId s, int r, int c) {
        Cell cell = getCell(s,r,c);
        recalcCell(s, r, c, cell);
        m_sheets[s].cells[{r,c}] = cell;
        return cell.cachedValue;
    }

    void recalcCell(SheetId s, int r, int c, Cell& cell) {
        if (cell.formula.isEmpty()) {
            cell.cachedValue = cell.rawValue;
            cell.dirty = false;
            return;
        }
        auto resolver = [this, s](const QString& ref) { return resolveCell(s, ref); };
        cell.cachedValue = m_formula.evaluate(cell.formula, resolver);
        if (!m_formula.lastError().isEmpty())
            emit formulaError(s, r, c, m_formula.lastError());
        cell.dirty = false;
    }

    void recalcDependents(SheetId s, int r, int c) {
        // Simple: recalc all dirty formula cells in the sheet
        for (auto it = m_sheets[s].cells.begin(); it != m_sheets[s].cells.end(); ++it) {
            Cell& cell = it.value();
            if (!cell.formula.isEmpty()) {
                auto deps = m_formula.dependencies(cell.formula);
                // Build ref string for changed cell
                QString ref;
                int col = it.key().second + 1;
                while (col > 0) { col--; ref.prepend(QChar('A' + col % 26)); col /= 26; }
                ref += QString::number(it.key().first + 1);
                if (deps.contains(ref)) {
                    cell.dirty = true;
                    recalcCell(s, it.key().first, it.key().second, cell);
                }
            }
        }
    }
};

#include "SpreadsheetCore.moc"

extern "C" CORE_API ISpreadsheetCore* createSpreadsheetCore() {
    return new SpreadsheetCoreImpl();
}
