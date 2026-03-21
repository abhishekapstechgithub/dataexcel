#include "SpreadsheetView.h"
#include <QKeyEvent>
#include <QContextMenuEvent>
#include <QMenu>
#include <QHeaderView>
#include <QScrollBar>
#include <QApplication>
#include <QClipboard>
#include <functional>

SpreadsheetView::SpreadsheetView(ISpreadsheetCore* core, SheetId sheet, QWidget* parent)
    : QTableView(parent), m_core(core)
{
    m_model = new SpreadsheetTableModel(core, sheet, this);
    QTableView::setModel(m_model);
    setupHeaders();

    connect(selectionModel(), &QItemSelectionModel::currentChanged,
            this, &SpreadsheetView::onCurrentChanged);

    // Perf: only paint visible cells
    setHorizontalScrollMode(ScrollPerPixel);
    setVerticalScrollMode(ScrollPerPixel);
    setSelectionMode(ExtendedSelection);
    setEditTriggers(DoubleClicked | AnyKeyPressed);
    setAlternatingRowColors(false);
    setGridStyle(Qt::SolidLine);
    setShowGrid(true);

    // Row/col default sizes
    horizontalHeader()->setDefaultSectionSize(80);
    verticalHeader()->setDefaultSectionSize(24);
    horizontalHeader()->setSectionsMovable(true);
    verticalHeader()->setSectionsMovable(false);

    // Enable row/col resize
    horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);
}

SpreadsheetView::~SpreadsheetView() {}

SpreadsheetTableModel* SpreadsheetView::model() const { return m_model; }

void SpreadsheetView::setupHeaders() {
    // Corner button style
    setStyleSheet(
        "QTableView { gridline-color: #d0d0d0; font-size: 12px; }"
        "QHeaderView::section { background: #f3f3f3; border: 1px solid #d0d0d0;"
        "  padding: 2px; font-size: 11px; }"
        "QHeaderView::section:selected { background: #217346; color: white; }"
    );
}

void SpreadsheetView::switchSheet(SheetId sheet) {
    m_model->setSheet(sheet);
}

int SpreadsheetView::currentRow() const { return currentIndex().row(); }
int SpreadsheetView::currentCol() const { return currentIndex().column(); }

CellFormat SpreadsheetView::currentCellFormat() const {
    return m_core->getCell(m_model->currentSheet(),
                           currentRow(), currentCol()).format;
}

void SpreadsheetView::applyToSelection(std::function<void(int,int)> fn) {
    for (auto& idx : selectedIndexes())
        fn(idx.row(), idx.column());
}

void SpreadsheetView::applyFormat(const CellFormat& fmt) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){ m_core->setCellFormat(s, r, c, fmt); });
}

void SpreadsheetView::applyBold(bool on) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.bold = on; m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyItalic(bool on) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.italic = on; m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyUnderline(bool on) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.underline = on; m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyTextColor(const QColor& col) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.textColor = col; m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyFillColor(const QColor& col) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.fillColor = col; m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyFontFamily(const QString& f) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.font.setFamily(f); m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyFontSize(int size) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.font.setPointSize(size); m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyHAlign(int a) {
    Qt::AlignmentFlag af = (a==1) ? Qt::AlignHCenter : (a==2) ? Qt::AlignRight : Qt::AlignLeft;
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.alignment = (cell.format.alignment & ~Qt::AlignHorizontal_Mask) | af;
        m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyVAlign(int a) {
    Qt::AlignmentFlag af = (a==0) ? Qt::AlignTop : (a==2) ? Qt::AlignBottom : Qt::AlignVCenter;
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.alignment = (cell.format.alignment & ~Qt::AlignVertical_Mask) | af;
        m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyNumberFormat(int fmt) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.numberFormat = fmt; m_core->setCellFormat(s,r,c,cell.format);
    });
}
void SpreadsheetView::applyWrapText(bool on) {
    SheetId s = m_model->currentSheet();
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(s,r,c);
        cell.format.wrapText = on; m_core->setCellFormat(s,r,c,cell.format);
    });
}

void SpreadsheetView::insertRow() {
    m_core->insertRow(m_model->currentSheet(), currentRow());
}
void SpreadsheetView::deleteRow() {
    m_core->deleteRow(m_model->currentSheet(), currentRow());
}
void SpreadsheetView::insertColumn() {
    m_core->insertColumn(m_model->currentSheet(), currentCol());
}
void SpreadsheetView::deleteColumn() {
    m_core->deleteColumn(m_model->currentSheet(), currentCol());
}
void SpreadsheetView::sortAsc() {
    auto sel = selectedIndexes();
    if (sel.isEmpty()) return;
    int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
    for (auto& i : sel) {
        r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
        c1=qMin(c1,i.column()); c2=qMax(c2,i.column());
    }
    m_core->sortRange(m_model->currentSheet(), r1,c1,r2,c2, c1, Qt::AscendingOrder);
}
void SpreadsheetView::sortDesc() {
    auto sel = selectedIndexes();
    if (sel.isEmpty()) return;
    int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
    for (auto& i : sel) {
        r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
        c1=qMin(c1,i.column()); c2=qMax(c2,i.column());
    }
    m_core->sortRange(m_model->currentSheet(), r1,c1,r2,c2, c1, Qt::DescendingOrder);
}

void SpreadsheetView::autoSum() {
    int r = currentRow(), c = currentCol();
    SheetId s = m_model->currentSheet();
    // Find contiguous non-empty cells above
    int startRow = r - 1;
    while (startRow >= 0 && m_core->getCell(s, startRow, c).rawValue.isValid())
        startRow--;
    startRow++;
    if (startRow == r) return;
    QString topRef    = SpreadsheetTableModel::columnLabel(c) + QString::number(startRow + 1);
    QString bottomRef = SpreadsheetTableModel::columnLabel(c) + QString::number(r);
    m_core->setCellFormula(s, r, c, QString("=SUM(%1:%2)").arg(topRef, bottomRef));
}

void SpreadsheetView::mergeSelected() {
    auto sel = selectedIndexes();
    if (sel.size() < 2) return;
    int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
    for (auto& i : sel) {
        r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
        c1=qMin(c1,i.column()); c2=qMax(c2,i.column());
    }
    m_core->mergeCells(m_model->currentSheet(), r1, c1, r2, c2);
}

void SpreadsheetView::keyPressEvent(QKeyEvent* e) {
    if (e->modifiers() == Qt::ControlModifier) {
        switch (e->key()) {
        case Qt::Key_C: {
            // Copy to clipboard
            QStringList rows;
            int lastRow = -1; QStringList row;
            for (auto& idx : selectedIndexes()) {
                if (lastRow != -1 && idx.row() != lastRow) {
                    rows << row.join('\t'); row.clear();
                }
                Cell c = m_core->getCell(m_model->currentSheet(), idx.row(), idx.column());
                row << (c.cachedValue.isValid() ? c.cachedValue.toString() : c.rawValue.toString());
                lastRow = idx.row();
            }
            if (!row.isEmpty()) rows << row.join('\t');
            QApplication::clipboard()->setText(rows.join('\n'));
            return;
        }
        case Qt::Key_V: {
            QString text = QApplication::clipboard()->text();
            QStringList lines = text.split('\n');
            int startRow = currentRow(), startCol = currentCol();
            SheetId s = m_model->currentSheet();
            for (int r = 0; r < lines.size(); ++r) {
                QStringList cells = lines[r].split('\t');
                for (int c = 0; c < cells.size(); ++c)
                    m_core->setCellValue(s, startRow+r, startCol+c, cells[c]);
            }
            return;
        }
        case Qt::Key_Z: m_core->undo(); return;
        case Qt::Key_Y: m_core->redo(); return;
        }
    }
    QTableView::keyPressEvent(e);
}

void SpreadsheetView::contextMenuEvent(QContextMenuEvent* e) {
    QMenu menu(this);
    menu.addAction("Cut",    this, [this]{ m_core->clearCell(m_model->currentSheet(), currentRow(), currentCol()); });
    menu.addAction("Copy",   this, [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress, Qt::Key_C, Qt::ControlModifier)); });
    menu.addAction("Paste",  this, [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress, Qt::Key_V, Qt::ControlModifier)); });
    menu.addSeparator();
    menu.addAction("Insert Row",    this, &SpreadsheetView::insertRow);
    menu.addAction("Delete Row",    this, &SpreadsheetView::deleteRow);
    menu.addAction("Insert Column", this, &SpreadsheetView::insertColumn);
    menu.addAction("Delete Column", this, &SpreadsheetView::deleteColumn);
    menu.addSeparator();
    menu.addAction("Sort A→Z", this, &SpreadsheetView::sortAsc);
    menu.addAction("Sort Z→A", this, &SpreadsheetView::sortDesc);
    menu.exec(e->globalPos());
}

void SpreadsheetView::onCurrentChanged(const QModelIndex& cur, const QModelIndex&) {
    if (!cur.isValid()) return;
    CellFormat fmt = m_core->getCell(m_model->currentSheet(), cur.row(), cur.column()).format;
    QString ref = SpreadsheetTableModel::columnLabel(cur.column()) +
                  QString::number(cur.row() + 1);
    emit selectionFormatChanged(fmt, ref);
}
