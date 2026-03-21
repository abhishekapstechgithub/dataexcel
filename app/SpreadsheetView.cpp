// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetView.cpp — Grid View (inspired by WPS griddrawer.dll)
//  Features:
//   - Excel-style header (selected col/row turns green)
//   - Ctrl+wheel zoom
//   - Multi-cell selection with copy/paste
//   - Auto-fit on header double-click
//   - Large virtual grid (200+ rows × 52+ cols)
// ═══════════════════════════════════════════════════════════════════════════════
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

SpreadsheetView::SpreadsheetView(ISpreadsheetCore* core, SheetId sheet, QWidget* parent)
    : QTableView(parent), m_core(core)
{
    m_model = new SpreadsheetTableModel(core, sheet, this);
    QTableView::setModel(m_model);

    // ── Performance settings ───────────────────────────────────────────────
    setHorizontalScrollMode(ScrollPerPixel);
    setVerticalScrollMode(ScrollPerPixel);
    setSelectionMode(ExtendedSelection);
    setEditTriggers(DoubleClicked | AnyKeyPressed);
    setAlternatingRowColors(false);
    setGridStyle(Qt::SolidLine);
    setShowGrid(true);
    setWordWrap(false);
    setTabKeyNavigation(false);  // We handle Tab ourselves

    // Minimum section sizes prevent columns collapsing to nothing
    horizontalHeader()->setMinimumSectionSize(8);
    verticalHeader()->setMinimumSectionSize(8);

    setupHeaders();

    connect(selectionModel(), &QItemSelectionModel::currentChanged,
            this, &SpreadsheetView::onCurrentChanged);
}

SpreadsheetView::~SpreadsheetView() {}

SpreadsheetTableModel* SpreadsheetView::model() const { return m_model; }

// ── Header & grid styling ─────────────────────────────────────────────────────
void SpreadsheetView::setupHeaders() {
    // The full WPS-style grid stylesheet
    setStyleSheet(
        // Grid
        "QTableView {"
        "  gridline-color: #d8d8d8;"
        "  font-family: 'Segoe UI', Calibri, Arial;"
        "  font-size: 11px;"
        "  background: #ffffff;"
        "  selection-background-color: #b8d9e8;"
        "  selection-color: #000000;"
        "  border: none;"
        "}"
        // Column/row headers — light grey like Excel/WPS
        "QHeaderView { font-family:'Segoe UI',Arial; }"
        "QHeaderView::section {"
        "  background: #f0f0f0;"
        "  border: none;"
        "  border-right: 1px solid #d0d0d0;"
        "  border-bottom: 1px solid #d0d0d0;"
        "  padding: 0 4px;"
        "  font-size: 11px;"
        "  color: #555;"
        "  font-family: 'Segoe UI', Arial;"
        "}"
        // Highlighted header (selected column/row) — WPS green
        "QHeaderView::section:checked {"
        "  background: #1e7145;"
        "  color: white;"
        "  font-weight: 600;"
        "}"
        // Select-all corner button
        "QAbstractButton {"
        "  background: #f0f0f0;"
        "  border-right: 1px solid #d0d0d0;"
        "  border-bottom: 1px solid #d0d0d0;"
        "}"
    );

    // Column headers
    horizontalHeader()->setDefaultSectionSize(m_baseColWidth);
    horizontalHeader()->setSectionsMovable(false);
    horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    horizontalHeader()->setHighlightSections(true);
    horizontalHeader()->setDefaultAlignment(Qt::AlignCenter | Qt::AlignVCenter);

    // Row headers — 50px wide showing row numbers right-aligned
    verticalHeader()->setDefaultSectionSize(m_baseRowHeight);
    verticalHeader()->setFixedWidth(50);
    verticalHeader()->setSectionsMovable(false);
    verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    verticalHeader()->setHighlightSections(true);
    verticalHeader()->setDefaultAlignment(Qt::AlignRight | Qt::AlignVCenter);
}

void SpreadsheetView::switchSheet(SheetId sheet) {
    m_model->setSheet(sheet);
    refreshSizes();
    // Scroll back to top-left on sheet switch (like WPS/Excel)
    scrollTo(m_model->index(0,0));
    setCurrentIndex(m_model->index(0,0));
}

// ── Zoom ──────────────────────────────────────────────────────────────────────
void SpreadsheetView::setZoomFactor(qreal factor) {
    m_zoomFactor = qBound(0.5, factor, 2.0);
    QFont f = font();
    f.setPointSizeF(qMax(6.0, 10.0 * m_zoomFactor));
    setFont(f);
    refreshSizes();
    viewport()->update();
}

void SpreadsheetView::refreshSizes() {
    int rh = qMax(14, (int)(m_baseRowHeight * m_zoomFactor));
    int cw = qMax(8,  (int)(m_baseColWidth  * m_zoomFactor));
    horizontalHeader()->setDefaultSectionSize(cw);
    verticalHeader()->setDefaultSectionSize(rh);
    QFont hf = horizontalHeader()->font();
    hf.setPointSizeF(qMax(7.0, 9.0 * m_zoomFactor));
    horizontalHeader()->setFont(hf);
    verticalHeader()->setFont(hf);
}

// ── Cell reference ─────────────────────────────────────────────────────────────
int SpreadsheetView::currentRow() const { return currentIndex().row(); }
int SpreadsheetView::currentCol() const { return currentIndex().column(); }

CellFormat SpreadsheetView::currentCellFormat() const {
    if (currentRow()<0||currentCol()<0) return {};
    return m_core->getCell(m_model->currentSheet(),currentRow(),currentCol()).format;
}

QString SpreadsheetView::selectionToRef() const {
    auto sel = selectedIndexes();
    if (sel.isEmpty()) {
        int r=currentRow(), c=currentCol();
        if (r<0||c<0) return "A1";
        return SpreadsheetTableModel::columnLabel(c)+QString::number(r+1);
    }
    int r1=sel.first().row(),r2=r1,c1=sel.first().column(),c2=c1;
    for (auto& i:sel){r1=qMin(r1,i.row());r2=qMax(r2,i.row());c1=qMin(c1,i.column());c2=qMax(c2,i.column());}
    QString ref = SpreadsheetTableModel::columnLabel(c1)+QString::number(r1+1);
    if (r1!=r2||c1!=c2) ref+=":"+SpreadsheetTableModel::columnLabel(c2)+QString::number(r2+1);
    return ref;
}

// ── Apply format helpers ──────────────────────────────────────────────────────
void SpreadsheetView::applyToSelection(std::function<void(int,int)> fn) {
    for (auto& idx : selectedIndexes()) fn(idx.row(),idx.column());
}

void SpreadsheetView::applyFormat(const CellFormat& fmt) {
    SheetId s=m_model->currentSheet();
    applyToSelection([&](int r,int c){ m_core->setCellFormat(s,r,c,fmt); });
}

#define APPLY(FIELD,VAL) { \
    SheetId s=m_model->currentSheet(); \
    applyToSelection([&](int r,int c){ \
        auto cell=m_core->getCell(s,r,c); cell.format.FIELD=VAL; \
        m_core->setCellFormat(s,r,c,cell.format); }); }

void SpreadsheetView::applyBold(bool v)          { APPLY(bold,v)        }
void SpreadsheetView::applyItalic(bool v)        { APPLY(italic,v)      }
void SpreadsheetView::applyUnderline(bool v)     { APPLY(underline,v)   }
void SpreadsheetView::applyWrapText(bool v)      { APPLY(wrapText,v)    }
void SpreadsheetView::applyTextColor(const QColor& c){ APPLY(textColor,c) }
void SpreadsheetView::applyFillColor(const QColor& c){ APPLY(fillColor,c) }
void SpreadsheetView::applyNumberFormat(int f)   { APPLY(numberFormat,f) }

void SpreadsheetView::applyFontFamily(const QString& fam) {
    SheetId s=m_model->currentSheet();
    applyToSelection([&](int r,int c){
        auto cell=m_core->getCell(s,r,c);
        cell.format.font.setFamily(fam); m_core->setCellFormat(s,r,c,cell.format); });
}
void SpreadsheetView::applyFontSize(int sz) {
    SheetId s=m_model->currentSheet();
    applyToSelection([&](int r,int c){
        auto cell=m_core->getCell(s,r,c);
        cell.format.font.setPointSize(sz); m_core->setCellFormat(s,r,c,cell.format); });
}
void SpreadsheetView::applyHAlign(int a) {
    Qt::Alignment al=(a==2)?Qt::AlignRight:(a==1)?Qt::AlignHCenter:Qt::AlignLeft;
    SheetId s=m_model->currentSheet();
    applyToSelection([&](int r,int c){
        auto cell=m_core->getCell(s,r,c);
        cell.format.alignment=(cell.format.alignment&Qt::AlignVertical_Mask)|al;
        m_core->setCellFormat(s,r,c,cell.format); });
}
void SpreadsheetView::applyVAlign(int a) {
    Qt::Alignment al=(a==0)?Qt::AlignTop:(a==2)?Qt::AlignBottom:Qt::AlignVCenter;
    SheetId s=m_model->currentSheet();
    applyToSelection([&](int r,int c){
        auto cell=m_core->getCell(s,r,c);
        cell.format.alignment=(cell.format.alignment&Qt::AlignHorizontal_Mask)|al;
        m_core->setCellFormat(s,r,c,cell.format); });
}

// ── Row/col ops ───────────────────────────────────────────────────────────────
void SpreadsheetView::insertRow()    { int r=currentRow(); if(r>=0) m_core->insertRow(m_model->currentSheet(),r); }
void SpreadsheetView::deleteRow()    { int r=currentRow(); if(r>=0) m_core->deleteRow(m_model->currentSheet(),r); }
void SpreadsheetView::insertColumn() { int c=currentCol(); if(c>=0) m_core->insertColumn(m_model->currentSheet(),c); }
void SpreadsheetView::deleteColumn() { int c=currentCol(); if(c>=0) m_core->deleteColumn(m_model->currentSheet(),c); }

void SpreadsheetView::sortAsc() {
    auto sel=selectedIndexes(); if(sel.isEmpty()) return;
    int r1=sel.first().row(),r2=r1,c1=sel.first().column(),c2=c1;
    for(auto&i:sel){r1=qMin(r1,i.row());r2=qMax(r2,i.row());c1=qMin(c1,i.column());c2=qMax(c2,i.column());}
    m_core->sortRange(m_model->currentSheet(),r1,c1,r2,c2,c1,Qt::AscendingOrder);
}
void SpreadsheetView::sortDesc() {
    auto sel=selectedIndexes(); if(sel.isEmpty()) return;
    int r1=sel.first().row(),r2=r1,c1=sel.first().column(),c2=c1;
    for(auto&i:sel){r1=qMin(r1,i.row());r2=qMax(r2,i.row());c1=qMin(c1,i.column());c2=qMax(c2,i.column());}
    m_core->sortRange(m_model->currentSheet(),r1,c1,r2,c2,c1,Qt::DescendingOrder);
}

void SpreadsheetView::autoSum() {
    int r=currentRow(),c=currentCol(); if(r<=0||c<0) return;
    SheetId s=m_model->currentSheet();
    int top=r-1;
    while(top>0){ Cell cell=m_core->getCell(s,top-1,c);
        if(!cell.rawValue.isNull()&&cell.rawValue.canConvert<double>()) top--;
        else break; }
    if(top<r){
        QString r1=SpreadsheetTableModel::columnLabel(c)+QString::number(top+1);
        QString r2=SpreadsheetTableModel::columnLabel(c)+QString::number(r);
        m_core->setCellFormula(s,r,c,"=SUM("+r1+":"+r2+")");
    }
}

void SpreadsheetView::mergeSelected() {
    auto sel=selectedIndexes(); if(sel.size()<2) return;
    int r1=sel.first().row(),r2=r1,c1=sel.first().column(),c2=c1;
    for(auto&i:sel){r1=qMin(r1,i.row());r2=qMax(r2,i.row());c1=qMin(c1,i.column());c2=qMax(c2,i.column());}
    m_core->mergeCells(m_model->currentSheet(),r1,c1,r2,c2);
}

// ── Keyboard ──────────────────────────────────────────────────────────────────
void SpreadsheetView::keyPressEvent(QKeyEvent* e) {
    // Copy
    if (e->matches(QKeySequence::Copy)) {
        auto sel=selectedIndexes(); if(sel.isEmpty()) return;
        int r1=sel.first().row(),r2=r1,c1=sel.first().column(),c2=c1;
        for(auto&i:sel){r1=qMin(r1,i.row());r2=qMax(r2,i.row());c1=qMin(c1,i.column());c2=qMax(c2,i.column());}
        SheetId s=m_model->currentSheet();
        QString text;
        for(int r=r1;r<=r2;r++){
            for(int c=c1;c<=c2;c++){
                if(c>c1) text+='\t';
                Cell cell=m_core->getCell(s,r,c);
                text+=(cell.cachedValue.isValid()?cell.cachedValue:cell.rawValue).toString();
            }
            text+='\n';
        }
        QApplication::clipboard()->setText(text);
        return;
    }
    // Paste
    if (e->matches(QKeySequence::Paste)) {
        QString text=QApplication::clipboard()->text(); if(text.isEmpty()) return;
        int sr=currentRow(),sc=currentCol(); if(sr<0||sc<0) return;
        SheetId s=m_model->currentSheet();
        QStringList rows=text.split('\n',Qt::SkipEmptyParts);
        for(int ri=0;ri<rows.size();ri++){
            QStringList cols=rows[ri].split('\t');
            for(int ci=0;ci<cols.size();ci++){
                QString val=cols[ci];
                if(val.startsWith('=')) m_core->setCellFormula(s,sr+ri,sc+ci,val);
                else m_core->setCellValue(s,sr+ri,sc+ci,val);
            }
        }
        return;
    }
    // Delete / Backspace → clear cell contents
    if (e->key()==Qt::Key_Delete||e->key()==Qt::Key_Backspace) {
        SheetId s=m_model->currentSheet();
        for(auto&idx:selectedIndexes()) m_core->clearCell(s,idx.row(),idx.column());
        return;
    }
    // Undo/Redo
    if (e->matches(QKeySequence::Undo)) { m_core->undo(); return; }
    if (e->matches(QKeySequence::Redo)) { m_core->redo(); return; }
    // Ctrl+Home → A1
    if (e->key()==Qt::Key_Home&&(e->modifiers()&Qt::ControlModifier)) {
        setCurrentIndex(m_model->index(0,0)); scrollToTop(); return; }
    // Ctrl+End → last cell
    if (e->key()==Qt::Key_End&&(e->modifiers()&Qt::ControlModifier)) {
        int r=qMax(0,m_core->rowCount(m_model->currentSheet())-1);
        int c=qMax(0,m_core->columnCount(m_model->currentSheet())-1);
        setCurrentIndex(m_model->index(r,c)); return; }
    // Tab → move right (WPS/Excel behavior)
    if (e->key()==Qt::Key_Tab) {
        int r=currentRow(),c=currentCol()+1;
        if(c>=m_model->columnCount()){c=0;r++;}
        setCurrentIndex(m_model->index(r,c)); return; }
    // Enter → move down
    if (e->key()==Qt::Key_Return&&!(e->modifiers()&Qt::ControlModifier)) {
        int r=currentRow()+1,c=currentCol();
        setCurrentIndex(m_model->index(qMin(r,m_model->rowCount()-1),c)); return; }
    QTableView::keyPressEvent(e);
}

// ── Context menu ──────────────────────────────────────────────────────────────
void SpreadsheetView::contextMenuEvent(QContextMenuEvent* e) {
    QMenu menu(this);
    menu.setStyleSheet(
        "QMenu { font-size:12px; font-family:'Segoe UI',Arial; background:white; "
        "border:1px solid #ddd; }"
        "QMenu::item { padding:5px 28px 5px 14px; }"
        "QMenu::item:selected { background:#e8f5ee; color:#1a6b35; }"
        "QMenu::separator { height:1px; background:#ececec; margin:3px 0; }"
    );
    menu.addAction("Cut",   [this]{ QApplication::sendEvent(this,new QKeyEvent(QEvent::KeyPress,Qt::Key_X,Qt::ControlModifier)); });
    menu.addAction("Copy",  [this]{ QApplication::sendEvent(this,new QKeyEvent(QEvent::KeyPress,Qt::Key_C,Qt::ControlModifier)); });
    menu.addAction("Paste", [this]{ QApplication::sendEvent(this,new QKeyEvent(QEvent::KeyPress,Qt::Key_V,Qt::ControlModifier)); });
    menu.addSeparator();
    menu.addAction("Insert Row",    this,&SpreadsheetView::insertRow);
    menu.addAction("Insert Column", this,&SpreadsheetView::insertColumn);
    menu.addAction("Delete Row",    this,&SpreadsheetView::deleteRow);
    menu.addAction("Delete Column", this,&SpreadsheetView::deleteColumn);
    menu.addSeparator();
    menu.addAction("Clear Contents",[this]{
        SheetId s=m_model->currentSheet();
        for(auto&idx:selectedIndexes()) m_core->clearCell(s,idx.row(),idx.column()); });
    menu.addSeparator();
    menu.addAction("Sort A → Z", this,&SpreadsheetView::sortAsc);
    menu.addAction("Sort Z → A", this,&SpreadsheetView::sortDesc);
    menu.exec(e->globalPos());
}

// ── Double-click on header → auto-fit column ─────────────────────────────────
void SpreadsheetView::mouseDoubleClickEvent(QMouseEvent* e) {
    QHeaderView* hh = horizontalHeader();
    if (e->pos().y() <= hh->height()) {
        int col = hh->logicalIndexAt(e->pos().x());
        if (col >= 0) { resizeColumnToContents(col); return; }
    }
    QHeaderView* vh = verticalHeader();
    if (e->pos().x() <= vh->width()) {
        int row = vh->logicalIndexAt(e->pos().y());
        if (row >= 0) { resizeRowToContents(row); return; }
    }
    QTableView::mouseDoubleClickEvent(e);
}

// ── Ctrl+Wheel = zoom ─────────────────────────────────────────────────────────
void SpreadsheetView::wheelEvent(QWheelEvent* e) {
    if (e->modifiers()&Qt::ControlModifier) {
        // Pass to parent for zoom handling
        e->ignore(); return;
    }
    QTableView::wheelEvent(e);
}

// ── Current cell changed ──────────────────────────────────────────────────────
void SpreadsheetView::onCurrentChanged(const QModelIndex& cur, const QModelIndex&) {
    if (!cur.isValid()) return;
    CellFormat fmt = m_core->getCell(m_model->currentSheet(),cur.row(),cur.column()).format;
    emit selectionFormatChanged(fmt, selectionToRef());
}
