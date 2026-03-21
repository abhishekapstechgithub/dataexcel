// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetEngine.cpp — Sparse cell storage + formula engine + dependency graph
//
//  Key design decisions:
//  1. Storage: std::unordered_map<CellAddress, EngineCell> — O(1) access,
//     only stores non-empty cells, handles 1M×16K grid with minimal memory.
//  2. Dependency graph: tracks which cells depend on which, enabling
//     targeted recalculation instead of full-sheet recalc.
//  3. Topological sort: evaluates formulas in dependency order to avoid
//     stale values when chains like A1→B1→C1 exist.
//  4. Undo/redo: stores snapshots of changed cells (not full sheet).
// ═══════════════════════════════════════════════════════════════════════════════
#include "../include/engine/SpreadsheetEngine.h"
#include "../../include/FormulaEngine/FormulaEngine.h"
#include "../../dll/FormulaEngine/Tokenizer.h"
#include <QHash>
#include <QSet>
#include <QMutex>
#include <QMutexLocker>
#include <QRegularExpression>
#include <algorithm>
#include <unordered_map>
#include <vector>
#include <functional>

// ── Merge range ───────────────────────────────────────────────────────────────
struct MergeRange {
    int r1,c1,r2,c2;
    bool contains(int r,int c) const {
        return r>=r1&&r<=r2&&c>=c1&&c<=c2;
    }
};

// ── Row/col metadata ──────────────────────────────────────────────────────────
struct RowInfo { int height{20}; bool hidden{false}; };
struct ColInfo { int width{80};  bool hidden{false}; };

// ── Sheet storage ─────────────────────────────────────────────────────────────
struct SheetData {
    QString name;
    std::unordered_map<CellAddress, EngineCell> cells;  // sparse storage
    QHash<int, RowInfo> rows;
    QHash<int, ColInfo> cols;
    QList<MergeRange>   merges;
    int maxRow{0}, maxCol{0};

    // Dependency graph: addr → set of addrs that use this cell
    QHash<CellAddress, QSet<CellAddress>> dependents;
    // Reverse: addr → set of addrs this cell depends on
    QHash<CellAddress, QSet<CellAddress>> dependencies;
};

// ── Undo command ──────────────────────────────────────────────────────────────
struct UndoCommand {
    std::function<void()> redo_fn;
    std::function<void()> undo_fn;
    QString description;
};

class UndoManager {
public:
    void push(UndoCommand cmd) {
        m_stack.erase(m_stack.begin()+m_pos, m_stack.end());
        m_stack.push_back(std::move(cmd));
        m_pos = (int)m_stack.size();
        m_stack.back().redo_fn(); // execute immediately
        if ((int)m_stack.size() > 200) {  // limit undo depth
            m_stack.erase(m_stack.begin());
            m_pos--;
        }
    }
    void undo() { if (canUndo()) m_stack[--m_pos].undo_fn(); }
    void redo() { if (canRedo()) m_stack[m_pos++].redo_fn(); }
    bool canUndo() const { return m_pos > 0; }
    bool canRedo() const { return m_pos < (int)m_stack.size(); }

private:
    std::vector<UndoCommand> m_stack;
    int m_pos{0};
};

// ═══════════════════════════════════════════════════════════════════════════════
//  Impl
// ═══════════════════════════════════════════════════════════════════════════════
struct SpreadsheetEngine::Impl {
    QHash<EngineSheetId, SheetData> sheets;
    QList<EngineSheetId>            sheetOrder;
    EngineSheetId                   nextId{0};
    UndoManager                     undo;
    FormulaEngine                   formula;
    mutable QMutex                  mutex;

    SheetData& sheet(EngineSheetId id) { return sheets[id]; }
    const SheetData& sheet(EngineSheetId id) const { return sheets.value(id); }

    // ── Cell resolution for formula evaluator ─────────────────────────────────
    QVariant resolveRef(EngineSheetId sid, const QString& ref) {
        // Support Sheet1!A1 cross-sheet references
        QString r = ref;
        EngineSheetId targetSheet = sid;
        if (ref.contains('!')) {
            QStringList parts = ref.split('!');
            for (auto it = sheets.begin(); it != sheets.end(); ++it) {
                if (it->name.compare(parts[0], Qt::CaseInsensitive) == 0) {
                    targetSheet = it.key();
                    break;
                }
            }
            r = parts.last();
        }
        CellAddress addr = CellAddress::fromRef(r);
        if (!addr.isValid()) return QVariant();
        auto& sd = sheet(targetSheet);
        auto cit = sd.cells.find(addr);
        if (cit == sd.cells.end()) return QVariant();
        const EngineCell& cell = cit->second;
        if (cell.dirty) {
            // Lazy evaluation: recalc this cell first
            const_cast<Impl*>(this)->recalcCell(targetSheet, addr.row, addr.col);
        }
        return cell.displayValue.isValid() ? cell.displayValue : cell.rawValue;
    }

    // ── Core recalc ───────────────────────────────────────────────────────────
    void recalcCell(EngineSheetId sid, int row, int col) {
        auto& sd = sheet(sid);
        CellAddress addr{row, col};
        auto it = sd.cells.find(addr);
        if (it == sd.cells.end()) return;

        EngineCell& cell = it->second;
        if (cell.formula.isEmpty()) {
            cell.displayValue = cell.rawValue;
            cell.dirty = false;
            return;
        }

        // Remove this cell from old dependency tracking
        for (auto& dep : sd.dependencies.value(addr))
            sd.dependents[dep].remove(addr);
        sd.dependencies.remove(addr);

        // Evaluate
        auto resolver = [this, sid](const QString& ref) -> QVariant {
            return resolveRef(sid, ref);
        };
        cell.displayValue = formula.evaluate(cell.formula, resolver);
        cell.dirty = false;

        // Rebuild dependency tracking
        auto deps = formula.dependencies(cell.formula);
        for (auto& depRef : deps) {
            CellAddress depAddr = CellAddress::fromRef(depRef);
            sd.dependencies[addr].insert(depAddr);
            sd.dependents[depAddr].insert(addr);
        }
    }

    // ── Topological sort for recalc order ─────────────────────────────────────
    // BFS from changed cell through dependents graph
    void recalcDependents(EngineSheetId sid, const CellAddress& changed) {
        auto& sd = sheet(sid);
        QSet<CellAddress> visited;
        QList<CellAddress> queue;

        // Start with direct dependents
        for (auto& dep : sd.dependents.value(changed))
            if (!visited.contains(dep)) { visited.insert(dep); queue.append(dep); }

        while (!queue.isEmpty()) {
            CellAddress addr = queue.takeFirst();
            auto it = sd.cells.find(addr);
            if (it != sd.cells.end() && !it->second.formula.isEmpty()) {
                it->second.dirty = true;
                recalcCell(sid, addr.row, addr.col);
            }
            // Propagate further
            for (auto& dep : sd.dependents.value(addr))
                if (!visited.contains(dep)) { visited.insert(dep); queue.append(dep); }
        }
    }

    // ── Update used dimensions ────────────────────────────────────────────────
    void updateDims(EngineSheetId sid, int row, int col) {
        auto& sd = sheet(sid);
        sd.maxRow = qMax(sd.maxRow, row);
        sd.maxCol = qMax(sd.maxCol, col);
    }

    void rebuildDims(EngineSheetId sid) {
        auto& sd = sheet(sid);
        sd.maxRow = 0; sd.maxCol = 0;
        for (auto& [addr, cell] : sd.cells) {
            if (!cell.isEmpty()) {
                sd.maxRow = qMax(sd.maxRow, addr.row);
                sd.maxCol = qMax(sd.maxCol, addr.col);
            }
        }
    }
};

// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetEngine — public API
// ═══════════════════════════════════════════════════════════════════════════════
SpreadsheetEngine::SpreadsheetEngine(QObject* parent)
    : QObject(parent), d(std::make_unique<Impl>())
{
    addSheet("Sheet1");
}

SpreadsheetEngine::~SpreadsheetEngine() = default;

// ── getCell ───────────────────────────────────────────────────────────────────
EngineCell SpreadsheetEngine::getCell(EngineSheetId sid, int row, int col) const {
    if (!d->sheets.contains(sid)) return {};
    auto& sd = d->sheet(sid);
    auto it = sd.cells.find(CellAddress{row,col});
    if (it == sd.cells.end()) return {};
    return it->second;
}

// ── setCellValue ──────────────────────────────────────────────────────────────
void SpreadsheetEngine::setCellValue(EngineSheetId sid, int row, int col,
                                      const QVariant& value)
{
    if (!d->sheets.contains(sid)) return;
    CellAddress addr{row,col};
    EngineCell oldCell = getCell(sid, row, col);
    EngineCell newCell = oldCell;
    newCell.formula.clear();
    newCell.rawValue     = value;
    newCell.displayValue = value;
    newCell.dirty        = false;
    newCell.type = value.isNull() ? CellType::Empty :
                   (value.typeId() == QMetaType::Bool)   ? CellType::Boolean :
                   (value.canConvert<double>()) ? CellType::Number : CellType::Text;

    d->undo.push({
        [this,sid,row,col,newCell,addr]() {
            d->sheet(sid).cells[addr] = newCell;
            d->updateDims(sid,row,col);
            d->recalcDependents(sid,addr);
            emit cellChanged(sid,row,col);
        },
        [this,sid,row,col,oldCell,addr]() {
            if (oldCell.isEmpty()) d->sheet(sid).cells.erase(addr);
            else d->sheet(sid).cells[addr] = oldCell;
            d->recalcDependents(sid,addr);
            emit cellChanged(sid,row,col);
        },
        "Set value"
    });
}

// ── setCellFormula ────────────────────────────────────────────────────────────
void SpreadsheetEngine::setCellFormula(EngineSheetId sid, int row, int col,
                                        const QString& formula)
{
    if (!d->sheets.contains(sid)) return;
    CellAddress addr{row,col};
    EngineCell oldCell = getCell(sid, row, col);
    EngineCell newCell = oldCell;
    newCell.formula      = formula;
    newCell.rawValue     = formula;
    newCell.type         = CellType::Formula;
    newCell.dirty        = true;

    // Evaluate immediately
    auto resolver = [this,sid](const QString& ref) -> QVariant {
        return d->resolveRef(sid, ref);
    };
    newCell.displayValue = d->formula.evaluate(formula, resolver);
    newCell.dirty        = false;

    d->undo.push({
        [this,sid,row,col,newCell,addr]() {
            d->sheet(sid).cells[addr] = newCell;
            d->updateDims(sid,row,col);
            d->recalcDependents(sid,addr);
            emit cellChanged(sid,row,col);
        },
        [this,sid,row,col,oldCell,addr]() {
            if (oldCell.isEmpty()) d->sheet(sid).cells.erase(addr);
            else d->sheet(sid).cells[addr] = oldCell;
            d->recalcDependents(sid,addr);
            emit cellChanged(sid,row,col);
        },
        "Set formula"
    });
}

// ── setCellFormat ──────────────────────────────────────────────────────────────
void SpreadsheetEngine::setCellFormat(EngineSheetId sid, int row, int col,
                                       const CellFormat& fmt)
{
    if (!d->sheets.contains(sid)) return;
    CellAddress addr{row,col};
    CellFormat oldFmt = d->sheet(sid).cells[addr].format;

    d->undo.push({
        [this,sid,row,col,fmt,addr]() {
            d->sheet(sid).cells[addr].format = fmt;
            emit cellChanged(sid,row,col);
        },
        [this,sid,row,col,oldFmt,addr]() {
            d->sheet(sid).cells[addr].format = oldFmt;
            emit cellChanged(sid,row,col);
        },
        "Format cell"
    });
}

// ── clearCell ─────────────────────────────────────────────────────────────────
void SpreadsheetEngine::clearCell(EngineSheetId sid, int row, int col) {
    if (!d->sheets.contains(sid)) return;
    CellAddress addr{row,col};
    auto it = d->sheet(sid).cells.find(addr);
    if (it == d->sheet(sid).cells.end()) return;

    EngineCell oldCell = it->second;
    d->undo.push({
        [this,sid,row,col,addr]() {
            d->sheet(sid).cells.erase(addr);
            d->recalcDependents(sid,addr);
            emit cellChanged(sid,row,col);
        },
        [this,sid,row,col,oldCell,addr]() {
            d->sheet(sid).cells[addr] = oldCell;
            d->recalcDependents(sid,addr);
            emit cellChanged(sid,row,col);
        },
        "Clear cell"
    });
}

void SpreadsheetEngine::clearRange(EngineSheetId sid, int r1, int c1, int r2, int c2) {
    for (int r=r1;r<=r2;r++) for (int c=c1;c<=c2;c++) clearCell(sid,r,c);
}

// ── bulkSetValue (no undo, for file loading) ──────────────────────────────────
void SpreadsheetEngine::bulkSetValue(EngineSheetId sid, int row, int col,
                                      const QVariant& value)
{
    if (!d->sheets.contains(sid)) return;
    CellAddress addr{row,col};
    EngineCell& cell = d->sheet(sid).cells[addr];
    cell.rawValue     = value;
    cell.displayValue = value;
    cell.formula.clear();
    cell.dirty = false;
    cell.type = value.isNull() ? CellType::Empty :
                (value.typeId() == QMetaType::Bool)   ? CellType::Boolean :
                (value.canConvert<double>()) ? CellType::Number : CellType::Text;
    d->updateDims(sid,row,col);
}

void SpreadsheetEngine::bulkFinalize(EngineSheetId sid) {
    recalcAll(sid);
    emit sheetChanged(sid);
}

// ── Dimensions ────────────────────────────────────────────────────────────────
int SpreadsheetEngine::usedRowCount(EngineSheetId sid) const {
    return d->sheets.contains(sid) ? d->sheet(sid).maxRow + 1 : 0;
}
int SpreadsheetEngine::usedColCount(EngineSheetId sid) const {
    return d->sheets.contains(sid) ? d->sheet(sid).maxCol + 1 : 0;
}
bool SpreadsheetEngine::cellExists(EngineSheetId sid, int row, int col) const {
    if (!d->sheets.contains(sid)) return false;
    return d->sheet(sid).cells.count(CellAddress{row,col}) > 0;
}

// ── Row/col metadata ──────────────────────────────────────────────────────────
int  SpreadsheetEngine::rowHeight(EngineSheetId sid, int row) const {
    return d->sheet(sid).rows.value(row, RowInfo{}).height; }
int  SpreadsheetEngine::colWidth(EngineSheetId sid, int col) const {
    return d->sheet(sid).cols.value(col, ColInfo{}).width; }
void SpreadsheetEngine::setRowHeight(EngineSheetId sid, int row, int h) {
    d->sheet(sid).rows[row].height = h; }
void SpreadsheetEngine::setColWidth(EngineSheetId sid, int col, int w) {
    d->sheet(sid).cols[col].width = w; }
bool SpreadsheetEngine::rowHidden(EngineSheetId sid, int row) const {
    return d->sheet(sid).rows.value(row).hidden; }
bool SpreadsheetEngine::colHidden(EngineSheetId sid, int col) const {
    return d->sheet(sid).cols.value(col).hidden; }

// ── Insert / delete rows ──────────────────────────────────────────────────────
void SpreadsheetEngine::insertRows(EngineSheetId sid, int before, int count) {
    if (!d->sheets.contains(sid)) return;
    auto& sd = d->sheet(sid);
    std::unordered_map<CellAddress,EngineCell> newCells;
    for (auto& [addr, cell] : sd.cells) {
        int r = addr.row;
        newCells[CellAddress{r>=before ? r+count : r, addr.col}] = cell;
    }
    sd.cells = std::move(newCells);
    sd.maxRow += count;
    emit sheetChanged(sid);
}

void SpreadsheetEngine::deleteRows(EngineSheetId sid, int row, int count) {
    if (!d->sheets.contains(sid)) return;
    auto& sd = d->sheet(sid);
    std::unordered_map<CellAddress,EngineCell> newCells;
    for (auto& [addr, cell] : sd.cells) {
        int r = addr.row;
        if (r >= row && r < row+count) continue; // deleted
        newCells[CellAddress{r>=row+count ? r-count : r, addr.col}] = cell;
    }
    sd.cells = std::move(newCells);
    sd.maxRow = qMax(0, sd.maxRow-count);
    emit sheetChanged(sid);
}

// ── Insert / delete cols ──────────────────────────────────────────────────────
void SpreadsheetEngine::insertCols(EngineSheetId sid, int before, int count) {
    if (!d->sheets.contains(sid)) return;
    auto& sd = d->sheet(sid);
    std::unordered_map<CellAddress,EngineCell> newCells;
    for (auto& [addr, cell] : sd.cells)
        newCells[CellAddress{addr.row, addr.col>=before ? addr.col+count : addr.col}] = cell;
    sd.cells = std::move(newCells);
    sd.maxCol += count;
    emit sheetChanged(sid);
}

void SpreadsheetEngine::deleteCols(EngineSheetId sid, int col, int count) {
    if (!d->sheets.contains(sid)) return;
    auto& sd = d->sheet(sid);
    std::unordered_map<CellAddress,EngineCell> newCells;
    for (auto& [addr, cell] : sd.cells) {
        int c = addr.col;
        if (c >= col && c < col+count) continue;
        newCells[CellAddress{addr.row, c>=col+count ? c-count : c}] = cell;
    }
    sd.cells = std::move(newCells);
    sd.maxCol = qMax(0, sd.maxCol-count);
    emit sheetChanged(sid);
}

// ── Merge ─────────────────────────────────────────────────────────────────────
void SpreadsheetEngine::mergeCells(EngineSheetId sid, int r1, int c1, int r2, int c2) {
    d->sheet(sid).merges.append({r1,c1,r2,c2});
    emit sheetChanged(sid);
}
void SpreadsheetEngine::unmergeCells(EngineSheetId sid, int r1, int c1) {
    auto& m = d->sheet(sid).merges;
    m.erase(std::remove_if(m.begin(),m.end(),[r1,c1](const MergeRange& mr){
        return mr.r1==r1&&mr.c1==c1;}), m.end());
    emit sheetChanged(sid);
}
bool SpreadsheetEngine::isMerged(EngineSheetId sid, int r, int c) const {
    for (auto& mr : d->sheet(sid).merges) if (mr.contains(r,c)) return true;
    return false;
}
CellAddress SpreadsheetEngine::mergeOrigin(EngineSheetId sid, int r, int c) const {
    for (auto& mr : d->sheet(sid).merges)
        if (mr.contains(r,c)) return {mr.r1,mr.c1};
    return {r,c};
}

// ── Sheet management ──────────────────────────────────────────────────────────
EngineSheetId SpreadsheetEngine::addSheet(const QString& name) {
    EngineSheetId id = d->nextId++;
    d->sheets[id].name = name.isEmpty() ? QString("Sheet%1").arg(d->sheetOrder.size()+1) : name;
    d->sheetOrder.append(id);
    return id;
}
void SpreadsheetEngine::removeSheet(EngineSheetId sid) {
    d->sheets.remove(sid); d->sheetOrder.removeAll(sid);
}
void SpreadsheetEngine::renameSheet(EngineSheetId sid, const QString& name) {
    if (d->sheets.contains(sid)) d->sheets[sid].name = name;
}
void SpreadsheetEngine::moveSheet(EngineSheetId sid, int newIdx) {
    d->sheetOrder.removeAll(sid);
    d->sheetOrder.insert(qBound(0,newIdx,d->sheetOrder.size()), sid);
}
QString SpreadsheetEngine::sheetName(EngineSheetId sid) const {
    return d->sheets.value(sid).name;
}
QList<EngineSheetId> SpreadsheetEngine::sheets() const { return d->sheetOrder; }
int SpreadsheetEngine::sheetIndex(EngineSheetId sid) const {
    return d->sheetOrder.indexOf(sid);
}

// ── Sort ──────────────────────────────────────────────────────────────────────
void SpreadsheetEngine::sortRange(EngineSheetId sid, int r1, int c1, int r2, int c2,
                                   int keyCol, bool ascending, bool /*hasHeader*/)
{
    if (!d->sheets.contains(sid)) return;
    auto& sd = d->sheet(sid);

    // Collect rows as vectors
    std::vector<std::vector<EngineCell>> rows;
    for (int r=r1;r<=r2;r++) {
        std::vector<EngineCell> row;
        for (int c=c1;c<=c2;c++) {
            auto it = sd.cells.find(CellAddress{r,c});
            row.push_back(it!=sd.cells.end() ? it->second : EngineCell{});
        }
        rows.push_back(std::move(row));
    }

    std::stable_sort(rows.begin(),rows.end(),
        [&](const std::vector<EngineCell>& a, const std::vector<EngineCell>& b){
            int ki = keyCol - c1;
            const EngineCell& ca = a[ki], &cb = b[ki];
            QVariant va = ca.displayValue.isValid() ? ca.displayValue : ca.rawValue;
            QVariant vb = cb.displayValue.isValid() ? cb.displayValue : cb.rawValue;
            bool na, nb;
            double da=va.toDouble(&na), db=vb.toDouble(&nb);
            bool less = (na&&nb) ? da<db : va.toString()<vb.toString();
            return ascending ? less : !less;
        });

    // Write back
    for (int r=r1;r<=r2;r++) {
        const auto& row = rows[r-r1];
        for (int c=c1;c<=c2;c++) {
            CellAddress addr{r,c};
            if (row[c-c1].isEmpty()) sd.cells.erase(addr);
            else sd.cells[addr] = row[c-c1];
        }
    }
    emit sheetChanged(sid);
}

// ── Undo/Redo ─────────────────────────────────────────────────────────────────
void SpreadsheetEngine::undo()        { d->undo.undo(); }
void SpreadsheetEngine::redo()        { d->undo.redo(); }
bool SpreadsheetEngine::canUndo() const { return d->undo.canUndo(); }
bool SpreadsheetEngine::canRedo() const { return d->undo.canRedo(); }

// ── Recalc ────────────────────────────────────────────────────────────────────
void SpreadsheetEngine::recalcAll(EngineSheetId sid) {
    if (!d->sheets.contains(sid)) return;
    auto& sd = d->sheet(sid);
    for (auto& [addr, cell] : sd.cells) {
        if (!cell.formula.isEmpty()) {
            cell.dirty = true;
            d->recalcCell(sid, addr.row, addr.col);
        }
    }
    emit recalcFinished(sid);
}

void SpreadsheetEngine::recalcCell(EngineSheetId sid, int row, int col) {
    d->recalcCell(sid,row,col);
    emit cellChanged(sid,row,col);
}

// ── Search ────────────────────────────────────────────────────────────────────
QList<SpreadsheetEngine::SearchResult>
SpreadsheetEngine::findAll(EngineSheetId sid, const QString& text,
                            bool matchCase, bool wholeCell) const
{
    QList<SearchResult> results;
    if (!d->sheets.contains(sid)) return results;
    Qt::CaseSensitivity cs = matchCase ? Qt::CaseSensitive : Qt::CaseInsensitive;
    for (auto& [addr, cell] : d->sheet(sid).cells) {
        QString val = (cell.displayValue.isValid() ? cell.displayValue : cell.rawValue).toString();
        bool match = wholeCell ? val.compare(text,cs)==0 : val.contains(text,cs);
        if (match) results.append({sid, addr.row, addr.col});
    }
    return results;
}
