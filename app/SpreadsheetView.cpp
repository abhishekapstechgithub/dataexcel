#include "SpreadsheetView.h"
#include <QKeyEvent>
#include <QContextMenuEvent>
#include <QMouseEvent>
#include <QWheelEvent>
#include <QMenu>
#include <QHeaderView>
#include <QScrollBar>
#include <QApplication>
#include <QClipboard>
#include <QRegularExpression>
#include <functional>

// ── Construction ──────────────────────────────────────────────────────────────
SpreadsheetView::SpreadsheetView(ISpreadsheetCore* core, SheetId sheet, QWidget* parent)
    : QTableView(parent), m_core(core)
{
    m_model = new SpreadsheetTableModel(core, sheet, this);
    QTableView::setModel(m_model);

    // Performance settings — only render visible cells
    setHorizontalScrollMode(ScrollPerPixel);
    setVerticalScrollMode(ScrollPerPixel);
    setSelectionMode(ExtendedSelection);
    setEditTriggers(DoubleClicked | AnyKeyPressed);
    setAlternatingRowColors(false);
    setGridStyle(Qt::SolidLine);
    setShowGrid(true);
    setWordWrap(false);

    // Minimum rows/cols visible — keeps viewport large enough to always show something
    horizontalHeader()->setMinimumSectionSize(24);
    verticalHeader()->setMinimumSectionSize(14);

    setupHeaders();

    connect(selectionModel(), &QItemSelectionModel::currentChanged,
            this, &SpreadsheetView::onCurrentChanged);
}

SpreadsheetView::~SpreadsheetView() {}

SpreadsheetTableModel* SpreadsheetView::model() const { return m_model; }

// ── Header setup ──────────────────────────────────────────────────────────────
void SpreadsheetView::setupHeaders() {
    // Excel-like header appearance
    setStyleSheet(
        // Grid
        "QTableView {"
        "  gridline-color: #d0d0d0;"
        "  font-size: 12px;"
        "  background: #ffffff;"
        "  selection-background-color: #c0d8f0;"
        "  selection-color: #000000;"
        "}"
        // Row/column headers
        "QHeaderView::section {"
        "  background: #f3f3f3;"
        "  border: none;"
        "  border-right: 1px solid #d0d0d0;"
        "  border-bottom: 1px solid #d0d0d0;"
        "  padding: 2px 4px;"
        "  font-size: 11px;"
        "  color: #444;"
        "}"
        // Headers for selected rows/cols turn green like Excel
        "QHeaderView::section:checked {"
        "  background: #1e7145;"
        "  color: white;"
        "  font-weight: bold;"
        "}"
        // Corner button (top-left select-all)
        "QAbstractButton {"
        "  background: #f3f3f3;"
        "  border: 1px solid #d0d0d0;"
        "}"
    );

    // Column headers: default width 80px, resizable, not movable
    horizontalHeader()->setDefaultSectionSize(m_baseColWidth);
    horizontalHeader()->setSectionsMovable(false);
    horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    horizontalHeader()->setHighlightSections(true);

    // Row headers: default height 24px, fixed width 50px for row numbers
    verticalHeader()->setDefaultSectionSize(m_baseRowHeight);
    verticalHeader()->setFixedWidth(50);
    verticalHeader()->setSectionsMovable(false);
    verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    verticalHeader()->setHighlightSections(true);
    verticalHeader()->setDefaultAlignment(Qt::AlignRight | Qt::AlignVCenter);
}

// ── Sheet switch ──────────────────────────────────────────────────────────────
void SpreadsheetView::switchSheet(SheetId sheet) {
    m_model->setSheet(sheet);
    // Restore default sizes after sheet change
    horizontalHeader()->setDefaultSectionSize(
        qMax(24, (int)(m_baseColWidth * m_zoomFactor)));
    verticalHeader()->setDefaultSectionSize(
        qMax(14, (int)(m_baseRowHeight * m_zoomFactor)));
}

// ── Zoom ──────────────────────────────────────────────────────────────────────
void SpreadsheetView::setZoomFactor(qreal factor) {
    m_zoomFactor = qBound(0.5, factor, 2.0);

    // Scale the view font
    QFont f = font();
    f.setPointSizeF(qMax(6.0, 10.0 * m_zoomFactor));
    setFont(f);

    // Scale row/column sizes
    int rowH = qMax(14, (int)(m_baseRowHeight * m_zoomFactor));
    int colW = qMax(24, (int)(m_baseColWidth  * m_zoomFactor));
    horizontalHeader()->setDefaultSectionSize(colW);
    verticalHeader()->setDefaultSectionSize(rowH);

    // Scale header font
    QFont hf = horizontalHeader()->font();
    hf.setPointSizeF(qMax(6.0, 9.0 * m_zoomFactor));
    horizontalHeader()->setFont(hf);
    verticalHeader()->setFont(hf);

    viewport()->update();
}

void SpreadsheetView::refreshSizes() {
    setZoomFactor(m_zoomFactor);
}

// ── Current cell ──────────────────────────────────────────────────────────────
int SpreadsheetView::currentRow() const { return currentIndex().row(); }
int SpreadsheetView::currentCol() const { return currentIndex().column(); }

CellFormat SpreadsheetView::currentCellFormat() const {
    if (currentRow() < 0 || currentCol() < 0) return {};
    return m_core->getCell(m_model->currentSheet(),
                           currentRow(), currentCol()).format;
}

// ── Cell reference string ─────────────────────────────────────────────────────
QString SpreadsheetView::selectionToRef() const {
    auto sel = selectedIndexes();
    if (sel.isEmpty()) {
        int r = currentRow(), c = currentCol();
        if (r < 0 || c < 0) return "A1";
        return SpreadsheetTableModel::columnLabel(c) + QString::number(r+1);
    }
    // Find bounding box of selection
    int r1 = sel.first().row(), r2 = r1;
    int c1 = sel.first().column(), c2 = c1;
    for (auto& idx : sel) {
        r1 = qMin(r1, idx.row());   r2 = qMax(r2, idx.row());
        c1 = qMin(c1, idx.column()); c2 = qMax(c2, idx.column());
    }
    QString ref = SpreadsheetTableModel::columnLabel(c1) + QString::number(r1+1);
    if (r1 != r2 || c1 != c2)
        ref += ":" + SpreadsheetTableModel::columnLabel(c2) + QString::number(r2+1);
    return ref;
}

// ── Apply format helpers ──────────────────────────────────────────────────────
void SpreadsheetView::applyToSelection(std::function<void(int,int)> fn) {
    for (auto& idx : selectedIndexes())
        fn(idx.row(), idx.column());
}

void SpreadsheetView::applyFormat(const CellFormat& fmt) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){ m_core->setCellFormat(s, r, c, fmt); });
}

#define APPLY_FORMAT_FIELD(FIELD, VALUE) \
    { SheetId s = m_model->currentSheet(); \
      applyToSelection([&](int r, int c){ \
          auto cell = m_core->getCell(s,r,c); \
          cell.format.FIELD = VALUE; \
          m_core->setCellFormat(s,r,c,cell.format); }); }

void SpreadsheetView::applyBold(bool on)          { APPLY_FORMAT_FIELD(bold,        on)  }
void SpreadsheetView::applyItalic(bool on)        { APPLY_FORMAT_FIELD(italic,      on)  }
void SpreadsheetView::applyUnderline(bool on)     { APPLY_FORMAT_FIELD(underline,   on)  }
void SpreadsheetView::applyWrapText(bool on)      { APPLY_FORMAT_FIELD(wrapText,    on)  }
void SpreadsheetView::applyTextColor(const QColor& c) { APPLY_FORMAT_FIELD(textColor, c) }
void SpreadsheetView::applyFillColor(const QColor& c) { APPLY_FORMAT_FIELD(fillColor, c) }
void SpreadsheetView::applyNumberFormat(int fmt)  { APPLY_FORMAT_FIELD(numberFormat, fmt)}

void SpreadsheetView::applyFontFamily(const QString& family) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        auto cell = m_core->getCell(s,r,c);
        cell.format.font.setFamily(family);
        m_core->setCellFormat(s,r,c,cell.format);
    });
}

void SpreadsheetView::applyFontSize(int size) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        auto cell = m_core->getCell(s,r,c);
        cell.format.font.setPointSize(size);
        m_core->setCellFormat(s,r,c,cell.format);
    });
}

void SpreadsheetView::applyHAlign(int a) {
    SheetId s = m_model->currentSheet();
    Qt::Alignment align = (a == 2) ? Qt::AlignRight :
                          (a == 1) ? Qt::AlignHCenter : Qt::AlignLeft;
    applyToSelection([&](int r, int c){
        auto cell = m_core->getCell(s,r,c);
        cell.format.alignment = (cell.format.alignment & Qt::AlignVertical_Mask) | align;
        m_core->setCellFormat(s,r,c,cell.format);
    });
}

void SpreadsheetView::applyVAlign(int a) {
    SheetId s = m_model->currentSheet();
    Qt::Alignment align = (a == 0) ? Qt::AlignTop :
                          (a == 2) ? Qt::AlignBottom : Qt::AlignVCenter;
    applyToSelection([&](int r, int c){
        auto cell = m_core->getCell(s,r,c);
        cell.format.alignment = (cell.format.alignment & Qt::AlignHorizontal_Mask) | align;
        m_core->setCellFormat(s,r,c,cell.format);
    });
}

// ── Row / column operations ───────────────────────────────────────────────────
void SpreadsheetView::insertRow() {
    int r = currentRow();
    if (r >= 0) m_core->insertRow(m_model->currentSheet(), r);
}

void SpreadsheetView::deleteRow() {
    int r = currentRow();
    if (r >= 0) m_core->deleteRow(m_model->currentSheet(), r);
}

void SpreadsheetView::insertColumn() {
    int c = currentCol();
    if (c >= 0) m_core->insertColumn(m_model->currentSheet(), c);
}

void SpreadsheetView::deleteColumn() {
    int c = currentCol();
    if (c >= 0) m_core->deleteColumn(m_model->currentSheet(), c);
}

void SpreadsheetView::sortAsc() {
    auto sel = selectedIndexes();
    if (sel.isEmpty()) return;
    int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
    for (auto& i:sel){ r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
                       c1=qMin(c1,i.column()); c2=qMax(c2,i.column()); }
    m_core->sortRange(m_model->currentSheet(), r1,c1,r2,c2, c1, Qt::AscendingOrder);
}

void SpreadsheetView::sortDesc() {
    auto sel = selectedIndexes();
    if (sel.isEmpty()) return;
    int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
    for (auto& i:sel){ r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
                       c1=qMin(c1,i.column()); c2=qMax(c2,i.column()); }
    m_core->sortRange(m_model->currentSheet(), r1,c1,r2,c2, c1, Qt::DescendingOrder);
}

void SpreadsheetView::autoSum() {
    int r = currentRow(), c = currentCol();
    if (r <= 0 || c < 0) return;
    // Find contiguous numbers above
    SheetId s = m_model->currentSheet();
    int topRow = r - 1;
    while (topRow > 0) {
        Cell cell = m_core->getCell(s, topRow-1, c);
        if (!cell.rawValue.isNull() && cell.rawValue.canConvert<double>()) topRow--;
        else break;
    }
    if (topRow < r) {
        QString ref1 = SpreadsheetTableModel::columnLabel(c) + QString::number(topRow+1);
        QString ref2 = SpreadsheetTableModel::columnLabel(c) + QString::number(r);
        m_core->setCellFormula(s, r, c, "=SUM(" + ref1 + ":" + ref2 + ")");
    }
}

void SpreadsheetView::mergeSelected() {
    auto sel = selectedIndexes();
    if (sel.size() < 2) return;
    int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
    for (auto& i:sel){ r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
                       c1=qMin(c1,i.column()); c2=qMax(c2,i.column()); }
    m_core->mergeCells(m_model->currentSheet(), r1,c1,r2,c2);
}

// ── Key handling ──────────────────────────────────────────────────────────────
void SpreadsheetView::keyPressEvent(QKeyEvent* e) {
    if (e->matches(QKeySequence::Copy)) {
        // Copy selection to clipboard as tab-separated text
        auto sel = selectedIndexes();
        if (sel.isEmpty()) return;
        int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
        for (auto& i:sel){ r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
                           c1=qMin(c1,i.column()); c2=qMax(c2,i.column()); }
        SheetId s = m_model->currentSheet();
        QString text;
        for (int r=r1; r<=r2; r++) {
            for (int c=c1; c<=c2; c++) {
                if (c > c1) text += '\t';
                Cell cell = m_core->getCell(s,r,c);
                text += (cell.cachedValue.isValid() ? cell.cachedValue : cell.rawValue).toString();
            }
            text += '\n';
        }
        QApplication::clipboard()->setText(text);
        return;
    }

    if (e->matches(QKeySequence::Paste)) {
        QString text = QApplication::clipboard()->text();
        if (text.isEmpty()) return;
        int startR = currentRow(), startC = currentCol();
        if (startR < 0 || startC < 0) return;
        SheetId s = m_model->currentSheet();
        QStringList rows = text.split('\n', Qt::SkipEmptyParts);
        for (int ri = 0; ri < rows.size(); ri++) {
            QStringList cols = rows[ri].split('\t');
            for (int ci = 0; ci < cols.size(); ci++) {
                QString val = cols[ci];
                if (val.startsWith('='))
                    m_core->setCellFormula(s, startR+ri, startC+ci, val);
                else
                    m_core->setCellValue(s, startR+ri, startC+ci, val);
            }
        }
        return;
    }

    if (e->key() == Qt::Key_Delete || e->key() == Qt::Key_Backspace) {
        SheetId s = m_model->currentSheet();
        for (auto& idx : selectedIndexes())
            m_core->clearCell(s, idx.row(), idx.column());
        return;
    }

    // Ctrl+Z / Ctrl+Y
    if (e->matches(QKeySequence::Undo)) { m_core->undo(); return; }
    if (e->matches(QKeySequence::Redo)) { m_core->redo(); return; }

    // Ctrl+Home: go to A1
    if (e->key() == Qt::Key_Home && e->modifiers() & Qt::ControlModifier) {
        setCurrentIndex(m_model->index(0, 0));
        return;
    }
    // Ctrl+End: go to last used cell
    if (e->key() == Qt::Key_End && e->modifiers() & Qt::ControlModifier) {
        int r = qMax(0, m_core->rowCount(m_model->currentSheet()) - 1);
        int c = qMax(0, m_core->columnCount(m_model->currentSheet()) - 1);
        setCurrentIndex(m_model->index(r, c));
        return;
    }

    // Tab: move right
    if (e->key() == Qt::Key_Tab) {
        int r = currentRow(), c = currentCol() + 1;
        if (c >= m_model->columnCount()) { c = 0; r++; }
        setCurrentIndex(m_model->index(r, c));
        return;
    }

    QTableView::keyPressEvent(e);
}

// ── Context menu ──────────────────────────────────────────────────────────────
void SpreadsheetView::contextMenuEvent(QContextMenuEvent* e) {
    QMenu menu(this);
    menu.setStyleSheet(
        "QMenu { font-size: 12px; }"
        "QMenu::item { padding: 5px 24px; }"
        "QMenu::item:selected { background: #e8f4ed; color: #1a6b35; }"
    );

    menu.addAction("Cut",   [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress, Qt::Key_X, Qt::ControlModifier)); });
    menu.addAction("Copy",  [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress, Qt::Key_C, Qt::ControlModifier)); });
    menu.addAction("Paste", [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress, Qt::Key_V, Qt::ControlModifier)); });
    menu.addSeparator();

    menu.addAction("Insert Row",    this, &SpreadsheetView::insertRow);
    menu.addAction("Delete Row",    this, &SpreadsheetView::deleteRow);
    menu.addAction("Insert Column", this, &SpreadsheetView::insertColumn);
    menu.addAction("Delete Column", this, &SpreadsheetView::deleteColumn);
    menu.addSeparator();

    menu.addAction("Clear Contents", [this]{
        SheetId s = m_model->currentSheet();
        for (auto& idx : selectedIndexes())
            m_core->clearCell(s, idx.row(), idx.column());
    });
    menu.addSeparator();
    menu.addAction("Sort A → Z",    this, &SpreadsheetView::sortAsc);
    menu.addAction("Sort Z → A",    this, &SpreadsheetView::sortDesc);

    menu.exec(e->globalPos());
}

// ── Double click: auto-fit column width ───────────────────────────────────────
void SpreadsheetView::mouseDoubleClickEvent(QMouseEvent* e) {
    // Check if click is on column header divider
    QHeaderView* hh = horizontalHeader();
    if (e->pos().y() <= hh->height()) {
        int logical = hh->logicalIndexAt(e->pos().x());
        if (logical >= 0) {
            resizeColumnToContents(logical);
            return;
        }
    }
    QTableView::mouseDoubleClickEvent(e);
}

// ── Wheel: Ctrl+wheel = zoom ──────────────────────────────────────────────────
void SpreadsheetView::wheelEvent(QWheelEvent* e) {
    if (e->modifiers() & Qt::ControlModifier) {
        // Zoom with Ctrl+scroll — emit to parent
        // Let the parent handle it by posting a synthetic event
        e->ignore();
        return;
    }
    QTableView::wheelEvent(e);
}

// ── Current changed → update formula bar ─────────────────────────────────────
void SpreadsheetView::onCurrentChanged(const QModelIndex& cur, const QModelIndex&) {
    if (!cur.isValid()) return;
    CellFormat fmt = m_core->getCell(m_model->currentSheet(),
                                     cur.row(), cur.column()).format;
    emit selectionFormatChanged(fmt, selectionToRef());
}
