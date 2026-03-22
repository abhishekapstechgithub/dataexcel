// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetView.cpp — Native Qt grid that always shows rows and columns
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
#include <QStyledItemDelegate>
#include <QPainter>
#include <functional>

// ── Custom delegate: draws cells like Excel / WPS ─────────────────────────────
class CellDelegate : public QStyledItemDelegate {
public:
    explicit CellDelegate(QObject* parent = nullptr) : QStyledItemDelegate(parent) {}

    void paint(QPainter* p, const QStyleOptionViewItem& opt,
               const QModelIndex& idx) const override
    {
        // Fill background
        QColor bg = idx.data(Qt::BackgroundRole).value<QBrush>().color();
        if (!bg.isValid() || bg == Qt::white || bg.alpha() == 0)
            bg = (opt.state & QStyle::State_Selected) ? QColor("#b8d9e8") : Qt::white;
        p->fillRect(opt.rect, bg);

        // Grid lines
        p->setPen(QPen(QColor("#d8dde4"), 0.5));
        p->drawLine(opt.rect.bottomLeft(), opt.rect.bottomRight());
        p->drawLine(opt.rect.topRight(), opt.rect.bottomRight());

        // Text
        QString text = idx.data(Qt::DisplayRole).toString();
        if (text.isEmpty()) return;

        QFont font = idx.data(Qt::FontRole).value<QFont>();
        if (!font.family().isEmpty()) p->setFont(font);
        else {
            QFont f = p->font();
            f.setFamily("Segoe UI");
            f.setPixelSize(13);
            p->setFont(f);
        }

        QColor fg = idx.data(Qt::ForegroundRole).value<QBrush>().color();
        if (!fg.isValid()) fg = QColor("#1a1d23");
        p->setPen(fg);

        Qt::Alignment align = (Qt::Alignment)idx.data(Qt::TextAlignmentRole).toInt();
        if (!align) {
            // Auto-detect: numbers right, text left
            bool ok; text.toDouble(&ok);
            align = ok ? (Qt::AlignRight | Qt::AlignVCenter) : (Qt::AlignLeft | Qt::AlignVCenter);
        }

        QRect textRect = opt.rect.adjusted(4, 1, -4, -1);
        p->drawText(textRect, align, text);
    }

    QSize sizeHint(const QStyleOptionViewItem&, const QModelIndex&) const override {
        return QSize(80, 22);
    }
};

// ── Construction ──────────────────────────────────────────────────────────────
SpreadsheetView::SpreadsheetView(ISpreadsheetCore* core, SheetId sheet, QWidget* parent)
    : QTableView(parent), m_core(core)
{
    m_model = new SpreadsheetTableModel(core, sheet, this);
    QTableView::setModel(m_model);

    // Custom delegate for clean cell rendering
    setItemDelegate(new CellDelegate(this));

    // Performance
    setHorizontalScrollMode(ScrollPerPixel);
    setVerticalScrollMode(ScrollPerPixel);
    setSelectionMode(ExtendedSelection);
    setSelectionBehavior(SelectItems);
    setEditTriggers(DoubleClicked | AnyKeyPressed);
    setAlternatingRowColors(false);
    setGridStyle(Qt::NoPen);       // We draw our own grid in the delegate
    setShowGrid(false);
    setWordWrap(false);
    setTabKeyNavigation(false);
    setCornerButtonEnabled(true);

    // Ensure the view fills its container
    setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    setMinimumSize(200, 200);

    setupHeaders();

    connect(selectionModel(), &QItemSelectionModel::currentChanged,
            this, &SpreadsheetView::onCurrentChanged);
}

SpreadsheetView::~SpreadsheetView() {}

SpreadsheetTableModel* SpreadsheetView::model() const { return m_model; }

// ── Stylesheet ────────────────────────────────────────────────────────────────
void SpreadsheetView::setupHeaders() {
    // Modern flat style — no ugly sunken borders
    setStyleSheet(
        "QTableView {"
        "  background: #ffffff;"
        "  selection-background-color: #b8d9e8;"
        "  selection-color: #1a1d23;"
        "  border: none;"
        "  outline: none;"
        "}"
        "QTableView::focus { outline: none; }"
        // Column headers
        "QHeaderView::section:horizontal {"
        "  background: #f2f4f7;"
        "  color: #555;"
        "  font-size: 11px;"
        "  font-weight: 600;"
        "  font-family: 'Segoe UI', Arial;"
        "  border: none;"
        "  border-right: 1px solid #d0d5de;"
        "  border-bottom: 2px solid #d0d5de;"
        "  padding: 0 4px;"
        "  min-height: 22px;"
        "}"
        "QHeaderView::section:horizontal:checked {"
        "  background: #1a7a45;"
        "  color: white;"
        "}"
        // Row headers
        "QHeaderView::section:vertical {"
        "  background: #f2f4f7;"
        "  color: #888;"
        "  font-size: 11px;"
        "  font-family: 'Segoe UI', Arial;"
        "  border: none;"
        "  border-right: 2px solid #d0d5de;"
        "  border-bottom: 1px solid #d8dde4;"
        "  padding: 0 6px 0 0;"
        "  text-align: right;"
        "  min-height: 22px;"
        "}"
        "QHeaderView::section:vertical:checked {"
        "  background: #1a7a45;"
        "  color: white;"
        "}"
        // Corner select-all button
        "QAbstractButton {"
        "  background: #f2f4f7;"
        "  border-right: 2px solid #d0d5de;"
        "  border-bottom: 2px solid #d0d5de;"
        "}"
        // Scrollbars
        "QScrollBar:horizontal, QScrollBar:vertical {"
        "  background: #f2f4f7; border: none;"
        "}"
        "QScrollBar:horizontal { height: 10px; }"
        "QScrollBar:vertical   { width:  10px; }"
        "QScrollBar::handle:horizontal, QScrollBar::handle:vertical {"
        "  background: #c8cdd6; border-radius: 5px;"
        "  min-width: 20px; min-height: 20px;"
        "  margin: 2px;"
        "}"
        "QScrollBar::handle:hover { background: #8c93a0; }"
        "QScrollBar::add-line, QScrollBar::sub-line { width:0; height:0; }"
    );

    // Column headers: default width and behaviour
    horizontalHeader()->setDefaultSectionSize(m_baseColWidth);
    horizontalHeader()->setSectionsMovable(false);
    horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    horizontalHeader()->setHighlightSections(true);
    horizontalHeader()->setDefaultAlignment(Qt::AlignCenter | Qt::AlignVCenter);
    horizontalHeader()->setMinimumSectionSize(10);
    horizontalHeader()->setFixedHeight(24);

    // Row headers: narrow, right-aligned numbers
    verticalHeader()->setDefaultSectionSize(m_baseRowHeight);
    verticalHeader()->setFixedWidth(46);
    verticalHeader()->setSectionsMovable(false);
    verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    verticalHeader()->setHighlightSections(true);
    verticalHeader()->setDefaultAlignment(Qt::AlignRight | Qt::AlignVCenter);
    verticalHeader()->setMinimumSectionSize(10);
}

void SpreadsheetView::switchSheet(SheetId sheet) {
    m_model->setSheet(sheet);
    scrollTo(m_model->index(0, 0));
    setCurrentIndex(m_model->index(0, 0));
}

// ── Zoom ──────────────────────────────────────────────────────────────────────
void SpreadsheetView::setZoomFactor(qreal factor) {
    m_zoomFactor = qBound(0.5, factor, 3.0);
    int colW = qRound(m_baseColWidth  * m_zoomFactor);
    int rowH = qRound(m_baseRowHeight * m_zoomFactor);
    horizontalHeader()->setDefaultSectionSize(colW);
    verticalHeader()->setDefaultSectionSize(rowH);

    QFont f;
    f.setFamily("Segoe UI");
    f.setPixelSize(qMax(9, qRound(13 * m_zoomFactor)));
    setFont(f);
    horizontalHeader()->setFont(f);
    verticalHeader()->setFont(f);
}

qreal SpreadsheetView::zoomFactor() const { return m_zoomFactor; }

// ── Current cell ──────────────────────────────────────────────────────────────
int SpreadsheetView::currentRow() const { return currentIndex().row(); }
int SpreadsheetView::currentCol() const { return currentIndex().column(); }

CellFormat SpreadsheetView::currentCellFormat() const {
    if (!currentIndex().isValid()) return {};
    return m_core->getCell(m_model->currentSheet(),
                           currentRow(), currentCol()).format;
}

// ── Apply helpers ─────────────────────────────────────────────────────────────
void SpreadsheetView::applyToSelection(std::function<void(int,int)> fn) {
    for (const auto& idx : selectedIndexes())
        fn(idx.row(), idx.column());
    viewport()->update();
}

void SpreadsheetView::applyBold(bool on) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.bold = on;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyItalic(bool on) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.italic = on;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyUnderline(bool on) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.underline = on;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyWrapText(bool on) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.wrapText = on;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyTextColor(const QColor& col) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.textColor = col;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyFillColor(const QColor& col) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.fillColor = col;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyFontFamily(const QString& f) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.font.setFamily(f);
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyFontSize(int sz) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.font.setPointSize(sz);
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyHAlign(int a) {
    Qt::AlignmentFlag af = (a==1)?Qt::AlignHCenter:(a==2)?Qt::AlignRight:Qt::AlignLeft;
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.alignment = (cell.format.alignment & ~Qt::AlignHorizontal_Mask) | af;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyVAlign(int a) {
    Qt::AlignmentFlag af = (a==0)?Qt::AlignTop:(a==2)?Qt::AlignBottom:Qt::AlignVCenter;
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.alignment = (cell.format.alignment & ~Qt::AlignVertical_Mask) | af;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyNumberFormat(int fmt) {
    applyToSelection([&](int r, int c){
        Cell cell = m_core->getCell(m_model->currentSheet(),r,c);
        cell.format.numberFormat = fmt;
        m_core->setCellFormat(m_model->currentSheet(),r,c,cell.format);
    });
}
void SpreadsheetView::applyFormat(const CellFormat& fmt) {
    applyToSelection([&](int r, int c){
        m_core->setCellFormat(m_model->currentSheet(),r,c,fmt);
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
void SpreadsheetView::sortAsc()  { doSort(Qt::AscendingOrder);  }
void SpreadsheetView::sortDesc() { doSort(Qt::DescendingOrder); }

void SpreadsheetView::doSort(Qt::SortOrder order) {
    auto sel = selectedIndexes();
    if (sel.isEmpty()) return;
    int r1=sel.first().row(), r2=r1, c1=sel.first().column(), c2=c1;
    for (auto& i : sel) {
        r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
        c1=qMin(c1,i.column()); c2=qMax(c2,i.column());
    }
    m_core->sortRange(m_model->currentSheet(), r1,c1,r2,c2, c1, order);
}

void SpreadsheetView::autoSum() {
    int r=currentRow(), c=currentCol();
    SheetId s=m_model->currentSheet();
    int startRow=r-1;
    while (startRow>=0 && m_core->getCell(s,startRow,c).rawValue.isValid()) startRow--;
    startRow++;
    if (startRow==r) return;
    QString top    = SpreadsheetTableModel::columnLabel(c)+QString::number(startRow+1);
    QString bottom = SpreadsheetTableModel::columnLabel(c)+QString::number(r);
    m_core->setCellFormula(s,r,c,QString("=SUM(%1:%2)").arg(top,bottom));
}

void SpreadsheetView::mergeSelected() {
    auto sel = selectedIndexes();
    if (sel.size()<2) return;
    int r1=sel.first().row(),r2=r1,c1=sel.first().column(),c2=c1;
    for (auto& i : sel) {
        r1=qMin(r1,i.row()); r2=qMax(r2,i.row());
        c1=qMin(c1,i.column()); c2=qMax(c2,i.column());
    }
    m_core->mergeCells(m_model->currentSheet(),r1,c1,r2,c2);
}

// ── Keyboard ──────────────────────────────────────────────────────────────────
void SpreadsheetView::keyPressEvent(QKeyEvent* e) {
    if (e->modifiers() == Qt::ControlModifier) {
        switch (e->key()) {
        case Qt::Key_C: {
            QStringList rows;
            int lastRow=-1; QStringList row;
            for (auto& idx : selectedIndexes()) {
                if (lastRow!=-1 && idx.row()!=lastRow) { rows<<row.join('\t'); row.clear(); }
                Cell c=m_core->getCell(m_model->currentSheet(),idx.row(),idx.column());
                row<<(c.cachedValue.isValid()?c.cachedValue.toString():c.rawValue.toString());
                lastRow=idx.row();
            }
            if (!row.isEmpty()) rows<<row.join('\t');
            QApplication::clipboard()->setText(rows.join('\n'));
            return;
        }
        case Qt::Key_V: {
            QString text=QApplication::clipboard()->text();
            QStringList lines=text.split('\n');
            int sr=currentRow(),sc=currentCol();
            SheetId s=m_model->currentSheet();
            for (int r=0;r<lines.size();++r) {
                QStringList cells=lines[r].split('\t');
                for (int c=0;c<cells.size();++c)
                    m_core->setCellValue(s,sr+r,sc+c,cells[c]);
            }
            return;
        }
        case Qt::Key_Z: m_core->undo(); return;
        case Qt::Key_Y: m_core->redo(); return;
        case Qt::Key_B: {
            bool b=!m_core->getCell(m_model->currentSheet(),currentRow(),currentCol()).format.bold;
            applyBold(b); emit boldToggled(b); return;
        }
        }
    }
    // Delete / Backspace
    if (e->key()==Qt::Key_Delete||e->key()==Qt::Key_Backspace) {
        for (auto& idx : selectedIndexes())
            m_core->clearCell(m_model->currentSheet(),idx.row(),idx.column());
        return;
    }
    QTableView::keyPressEvent(e);
}

void SpreadsheetView::contextMenuEvent(QContextMenuEvent* e) {
    QMenu menu(this);
    menu.setStyleSheet(
        "QMenu{background:white;border:1px solid #dde1e7;border-radius:5px;padding:4px 0;font-family:'Segoe UI';font-size:12px;}"
        "QMenu::item{padding:7px 20px 7px 14px;color:#1a1d23;}"
        "QMenu::item:selected{background:#e8f5ee;color:#1a7a45;}"
        "QMenu::separator{height:1px;background:#dde1e7;margin:4px 0;}"
    );
    menu.addAction("Cut",   this,[this]{ m_core->clearCell(m_model->currentSheet(),currentRow(),currentCol()); });
    menu.addAction("Copy",  this,[this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress,Qt::Key_C,Qt::ControlModifier)); });
    menu.addAction("Paste", this,[this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress,Qt::Key_V,Qt::ControlModifier)); });
    menu.addSeparator();
    menu.addAction("Insert Row",    this,&SpreadsheetView::insertRow);
    menu.addAction("Delete Row",    this,&SpreadsheetView::deleteRow);
    menu.addAction("Insert Column", this,&SpreadsheetView::insertColumn);
    menu.addAction("Delete Column", this,&SpreadsheetView::deleteColumn);
    menu.addSeparator();
    menu.addAction("Sort A → Z", this,&SpreadsheetView::sortAsc);
    menu.addAction("Sort Z → A", this,&SpreadsheetView::sortDesc);
    menu.addSeparator();
    menu.addAction("Undo", m_core,&ISpreadsheetCore::undo);
    menu.addAction("Redo", m_core,&ISpreadsheetCore::redo);
    menu.exec(e->globalPos());
}

void SpreadsheetView::wheelEvent(QWheelEvent* e) {
    if (e->modifiers() & Qt::ControlModifier) {
        qreal delta = e->angleDelta().y() > 0 ? 0.1 : -0.1;
        setZoomFactor(m_zoomFactor + delta);
        emit zoomChanged(qRound(m_zoomFactor * 100));
        e->accept();
        return;
    }
    QTableView::wheelEvent(e);
}

void SpreadsheetView::onCurrentChanged(const QModelIndex& cur, const QModelIndex&) {
    if (!cur.isValid()) return;
    CellFormat fmt = m_core->getCell(m_model->currentSheet(),cur.row(),cur.column()).format;
    QString ref = SpreadsheetTableModel::columnLabel(cur.column()) + QString::number(cur.row()+1);
    emit selectionFormatChanged(fmt, ref);
}
