// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetView.cpp — Virtual grid renderer for 100GB spreadsheet support
// ═══════════════════════════════════════════════════════════════════════════════
#include "SpreadsheetView.h"
#include <SpreadsheetCore/ISpreadsheetCore.h>
#include <QPainter>
#include <QPainterPath>
#include <QScrollBar>
#include <QLineEdit>
#include <QKeyEvent>
#include <QMouseEvent>
#include <QWheelEvent>
#include <QContextMenuEvent>
#include <QMenu>
#include <QAction>
#include <QClipboard>
#include <QApplication>
#include <QMimeData>
#include <QGuiApplication>
#include <QFont>
#include <QFontMetrics>
#include <QResizeEvent>
#include <QRegularExpression>
#include <QFrame>
#include <QPaintEvent>
#include <climits>
#include <cmath>
#include <algorithm>

// ── Column label helper ──────────────────────────────────────────────────────
QString SpreadsheetView::colLabel(int col)
{
    // col is 0-based → A=0, Z=25, AA=26, ...
    QString result;
    col++;
    while (col > 0) {
        col--;
        result.prepend(QChar('A' + col % 26));
        col /= 26;
    }
    return result;
}

int SpreadsheetView::colIndex(const QString& label)
{
    int result = 0;
    for (QChar c : label.toUpper()) {
        if (c < 'A' || c > 'Z') break;
        result = result * 26 + (c.unicode() - 'A' + 1);
    }
    return result - 1;  // 0-based
}

// ═════════════════════════════════════════════════════════════════════════════
SpreadsheetView::SpreadsheetView(QWidget* parent)
    : QAbstractScrollArea(parent)
{
    setFrameShape(QFrame::NoFrame);
    setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOn);
    setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOn);

    viewport()->setMouseTracking(true);
    viewport()->setCursor(Qt::ArrowCursor);
    viewport()->setBackgroundRole(QPalette::Base);
    viewport()->setAttribute(Qt::WA_OpaquePaintEvent);

    // Inline editor
    m_editor = new QLineEdit(viewport());
    m_editor->hide();
    m_editor->setFrame(false);
    m_editor->setFont(QFont("Calibri", 10));
    m_editor->setStyleSheet(
        "QLineEdit { background: white; border: 2px solid #217346; padding: 1px 2px; }"
    );
    connect(m_editor, &QLineEdit::returnPressed, this, &SpreadsheetView::commitEdit);

    // Scrollbar range
    updateScrollBars();
}

SpreadsheetView::~SpreadsheetView() = default;

// ── Data source ───────────────────────────────────────────────────────────────
void SpreadsheetView::setCore(ISpreadsheetCore* core)
{
    if (m_core) {
        disconnect(m_core, nullptr, this, nullptr);
    }
    m_core = core;
    if (m_core) {
        connect(m_core, &ISpreadsheetCore::cellChanged,
                this, &SpreadsheetView::refreshCell);
        connect(m_core, &ISpreadsheetCore::sheetChanged,
                this, [this](int){ refresh(); });
    }
    refresh();
}

void SpreadsheetView::setSheet(int sheetId)
{
    m_sheetId = sheetId;
    m_firstRow = 0;
    m_firstCol = 0;
    m_subpixelX = 0;
    m_subpixelY = 0;
    setCurrentCell(0, 0);
    updateScrollBars();
    refresh();
}

void SpreadsheetView::setTotalRows(qint64 rows)
{
    m_totalRows = rows;
    updateScrollBars();
}

void SpreadsheetView::setTotalCols(int cols)
{
    m_totalCols = cols;
    updateScrollBars();
}

// ── Column / Row sizes ────────────────────────────────────────────────────────
int SpreadsheetView::colWidth(int col) const
{
    if (m_core) {
        auto meta = m_core->getColMeta(m_sheetId, col);
        if (meta.width > 0) return (int)(meta.width * m_scale);
    }
    return m_colWidths.value(col, (int)(m_defaultColW * m_scale));
}

int SpreadsheetView::rowHeight(int row) const
{
    if (m_core) {
        auto meta = m_core->getRowMeta(m_sheetId, row);
        if (meta.height > 0) return (int)(meta.height * m_scale);
    }
    return m_rowHeights.value(row, (int)(m_defaultRowH * m_scale));
}

void SpreadsheetView::setColWidth(int col, int width)
{
    m_colWidths[col] = width;
    if (m_core) {
        ColumnMeta m = m_core->getColMeta(m_sheetId, col);
        m.width = (int)(width / m_scale);
        m_core->setColMeta(m_sheetId, col, m);
    }
    emit columnResized(col, width);
    updateScrollBars();
    refresh();
}

void SpreadsheetView::setRowHeight(int row, int height)
{
    m_rowHeights[row] = height;
    if (m_core) {
        RowMeta m = m_core->getRowMeta(m_sheetId, row);
        m.height = (int)(height / m_scale);
        m_core->setRowMeta(m_sheetId, row, m);
    }
    emit rowResized(row, height);
    refresh();
}

// ── Freeze panes ──────────────────────────────────────────────────────────────
void SpreadsheetView::setFreezeRow(int row) { m_freezeRow = row; refresh(); }
void SpreadsheetView::setFreezeCol(int col) { m_freezeCol = col; refresh(); }

// ── Zoom ──────────────────────────────────────────────────────────────────────
void SpreadsheetView::setZoom(int percent)
{
    m_zoom  = std::clamp(percent, 50, 400);
    m_scale = m_zoom / 100.0;
    updateScrollBars();
    refresh();
}

// ── Coordinate helpers ────────────────────────────────────────────────────────
int SpreadsheetView::yOfRow(int row) const
{
    // Pixel Y of top edge of row, relative to viewport
    int y = HEADER_ROW_H;
    // Frozen rows always at top
    for (int r = 0; r < m_freezeRow && r < row; ++r)
        y += rowHeight(r);

    if (row < m_freezeRow)
        return y - (m_freezeRow - row) * (int)(m_defaultRowH * m_scale); // rough

    // Scrolled rows
    qint64 start = std::max((qint64)m_freezeRow, m_firstRow);
    for (qint64 r = start; r < row; ++r)
        y += rowHeight((int)r);
    y -= m_subpixelY;
    return y;
}

int SpreadsheetView::xOfCol(int col) const
{
    int x = HEADER_COL_W;
    for (int c = 0; c < m_freezeCol && c < col; ++c)
        x += colWidth(c);

    if (col < m_freezeCol)
        return x - (m_freezeCol - col) * (int)(m_defaultColW * m_scale);

    int start = std::max(m_freezeCol, m_firstCol);
    for (int c = start; c < col; ++c)
        x += colWidth(c);
    x -= m_subpixelX;
    return x;
}

QRect SpreadsheetView::cellRect(int row, int col) const
{
    int x = xOfCol(col);
    int y = yOfRow(row);
    return QRect(x, y, colWidth(col), rowHeight(row));
}

int SpreadsheetView::rowAtY(int y) const
{
    if (y < HEADER_ROW_H) return -1;
    int py = HEADER_ROW_H - m_subpixelY;
    for (qint64 r = m_firstRow; r < m_totalRows; ++r) {
        int h = rowHeight((int)r);
        if (y >= py && y < py + h) return (int)r;
        py += h;
        if (py > viewport()->height()) break;
    }
    return -1;
}

int SpreadsheetView::colAtX(int x) const
{
    if (x < HEADER_COL_W) return -1;
    int px = HEADER_COL_W - m_subpixelX;
    for (int c = m_firstCol; c < m_totalCols; ++c) {
        int w = colWidth(c);
        if (x >= px && x < px + w) return c;
        px += w;
        if (px > viewport()->width()) break;
    }
    return -1;
}

// ── Visible range ─────────────────────────────────────────────────────────────
int SpreadsheetView::firstVisibleRow() const { return (int)m_firstRow; }

int SpreadsheetView::lastVisibleRow() const
{
    int h = HEADER_ROW_H - m_subpixelY;
    for (qint64 r = m_firstRow; r < m_totalRows; ++r) {
        h += rowHeight((int)r);
        if (h >= viewport()->height()) return (int)r;
    }
    return (int)m_totalRows - 1;
}

int SpreadsheetView::firstVisibleCol() const { return m_firstCol; }

int SpreadsheetView::lastVisibleCol() const
{
    int w = HEADER_COL_W - m_subpixelX;
    for (int c = m_firstCol; c < m_totalCols; ++c) {
        w += colWidth(c);
        if (w >= viewport()->width()) return c;
    }
    return m_totalCols - 1;
}

// ── Virtual scrollbar ─────────────────────────────────────────────────────────
static constexpr int VSCROLL_MAX = 100000;  // scrollbar maximum

qint64 SpreadsheetView::scrollValueToRow(int value) const
{
    if (m_totalRows <= 0) return 0;
    return (qint64)((double)value / VSCROLL_MAX * m_totalRows);
}

int SpreadsheetView::rowToScrollValue(qint64 row) const
{
    if (m_totalRows <= 0) return 0;
    return (int)((double)row / m_totalRows * VSCROLL_MAX);
}

void SpreadsheetView::updateScrollBars()
{
    // Vertical: virtual scroll
    verticalScrollBar()->setRange(0, VSCROLL_MAX);
    verticalScrollBar()->setPageStep(std::max(1, (int)((double)viewport()->height() / m_defaultRowH / m_totalRows * VSCROLL_MAX)));
    verticalScrollBar()->setSingleStep(std::max(1, VSCROLL_MAX / (int)m_totalRows));

    // Horizontal: pixel-based (columns are not 1 billion wide)
    int totalPixelW = HEADER_COL_W;
    for (int c = 0; c < m_totalCols; ++c)
        totalPixelW += colWidth(c);
    horizontalScrollBar()->setRange(0, std::max(0, totalPixelW - viewport()->width()));
    horizontalScrollBar()->setPageStep(viewport()->width());
    horizontalScrollBar()->setSingleStep(colWidth(m_firstCol));
}

void SpreadsheetView::scrollContentsBy(int dx, int dy)
{
    Q_UNUSED(dx);

    // Vertical: map scrollbar → row
    if (dy != 0) {
        int vval = verticalScrollBar()->value();
        m_firstRow = scrollValueToRow(vval);
        m_firstRow = std::clamp(m_firstRow, (qint64)0, m_totalRows - 1);
        m_subpixelY = 0;
    }

    // Horizontal: pixel-based
    int hval = horizontalScrollBar()->value();
    // Find first column that starts at hval
    int px = 0;
    m_firstCol = 0;
    m_subpixelX = 0;
    for (int c = 0; c < m_totalCols; ++c) {
        int w = colWidth(c);
        if (px + w > hval) {
            m_firstCol = c;
            m_subpixelX = hval - px;
            break;
        }
        px += w;
    }

    viewport()->update();
    if (m_editing) placeEditor();
}

// ── paintEvent ────────────────────────────────────────────────────────────────
void SpreadsheetView::paintEvent(QPaintEvent* e)
{
    QPainter p(viewport());
    p.setRenderHint(QPainter::TextAntialiasing);
    QRect clip = e->rect();

    // Background
    p.fillRect(clip, Qt::white);

    drawGridLines(p, clip);
    drawCells(p, clip);
    drawHeaders(p, clip);
    drawSelection(p);

    // Corner cell (top-left intersection of headers)
    p.fillRect(0, 0, HEADER_COL_W, HEADER_ROW_H, QColor(0xF2F2F2));
    p.setPen(QColor(0xD0D0D0));
    p.drawRect(0, 0, HEADER_COL_W - 1, HEADER_ROW_H - 1);

    // Clip line between header and grid
    p.setPen(QPen(QColor(0xB0B0B0), 1));
    p.drawLine(0, HEADER_ROW_H - 1, viewport()->width(), HEADER_ROW_H - 1);
    p.drawLine(HEADER_COL_W - 1, 0, HEADER_COL_W - 1, viewport()->height());
}

void SpreadsheetView::drawGridLines(QPainter& p, const QRect& clip)
{
    p.setPen(QPen(m_gridColor, 1));

    int lastR = lastVisibleRow();
    int lastC = lastVisibleCol();

    // Horizontal lines
    int y = HEADER_ROW_H - m_subpixelY;
    for (qint64 r = m_firstRow; r <= lastR + 1; ++r) {
        if (y >= clip.top() && y <= clip.bottom())
            p.drawLine(HEADER_COL_W, y, viewport()->width(), y);
        y += rowHeight((int)r);
        if (y > viewport()->height()) break;
    }

    // Vertical lines
    int x = HEADER_COL_W - m_subpixelX;
    for (int c = m_firstCol; c <= lastC + 1; ++c) {
        if (x >= clip.left() && x <= clip.right())
            p.drawLine(x, HEADER_ROW_H, x, viewport()->height());
        x += colWidth(c);
        if (x > viewport()->width()) break;
    }
}

void SpreadsheetView::drawCells(QPainter& p, const QRect& clip)
{
    if (!m_core) return;

    int lastR = lastVisibleRow();
    int lastC = lastVisibleCol();

    QFont baseFont("Calibri", (int)(10 * m_scale));

    for (qint64 r = m_firstRow; r <= lastR; ++r) {
        for (int c = m_firstCol; c <= lastC; ++c) {
            QRect cr = cellRect((int)r, c);
            if (!cr.intersects(clip)) continue;

            Cell cell = m_core->getCell(m_sheetId, (int)r, c);
            if (cell.rawValue.isNull() && cell.cachedValue.isNull()) continue;

            const CellFormat& fmt = cell.format;

            // Background
            QColor bg = fmt.fillColor;
            if (bg != Qt::white)
                p.fillRect(cr.adjusted(1, 1, 0, 0), bg);

            // Text
            QString text = cell.cachedValue.isNull()
                         ? cell.rawValue.toString()
                         : cell.cachedValue.toString();

            if (text.isEmpty()) continue;

            QFont f = baseFont;
            f.setBold(fmt.bold);
            f.setItalic(fmt.italic);
            f.setUnderline(fmt.underline);
            if (!fmt.font.family().isEmpty())
                f.setFamily(fmt.font.family());

            p.setFont(f);
            p.setPen(fmt.textColor);

            Qt::Alignment align = fmt.alignment;
            // Default numeric alignment = right
            bool isNum = cell.cachedValue.typeId() == QMetaType::Double ||
                         cell.cachedValue.typeId() == QMetaType::LongLong ||
                         cell.cachedValue.typeId() == QMetaType::Int;
            if ((align & Qt::AlignHorizontal_Mask) == 0) {
                align |= isNum ? Qt::AlignRight : Qt::AlignLeft;
            }
            if ((align & Qt::AlignVertical_Mask) == 0)
                align |= Qt::AlignVCenter;

            QRect textRect = cr.adjusted(3, 1, -3, -1);
            if (fmt.wrapText) {
                p.drawText(textRect, align | Qt::TextWordWrap, text);
            } else {
                p.drawText(textRect, align, text);
            }
        }
    }
}

void SpreadsheetView::drawHeaders(QPainter& p, const QRect& clip)
{
    int lastR = lastVisibleRow();
    int lastC = lastVisibleCol();

    QFont hdrFont("Segoe UI", (int)(9 * m_scale));
    p.setFont(hdrFont);

    // ── Column headers ────────────────────────────────────────────────────────
    {
        int x = HEADER_COL_W - m_subpixelX;
        for (int c = m_firstCol; c <= lastC; ++c) {
            int w = colWidth(c);
            QRect hr(x, 0, w, HEADER_ROW_H);
            bool active = (c >= m_selection.c1 && c <= m_selection.c2);

            p.fillRect(hr, active ? m_activeHeaderBg : m_headerBg);
            p.setPen(QColor(0xC8C8C8));
            p.drawRect(hr.adjusted(0, 0, -1, -1));

            p.setPen(active ? m_activeHeaderText : m_headerText);
            p.drawText(hr, Qt::AlignCenter, colLabel(c));

            x += w;
            if (x > viewport()->width()) break;
        }
    }

    // ── Row headers ───────────────────────────────────────────────────────────
    {
        int y = HEADER_ROW_H - m_subpixelY;
        for (qint64 r = m_firstRow; r <= lastR; ++r) {
            int h = rowHeight((int)r);
            QRect hr(0, y, HEADER_COL_W, h);
            bool active = (r >= m_selection.r1 && r <= m_selection.r2);

            p.fillRect(hr, active ? m_activeHeaderBg : m_headerBg);
            p.setPen(QColor(0xC8C8C8));
            p.drawRect(hr.adjusted(0, 0, -1, -1));

            p.setPen(active ? m_activeHeaderText : m_headerText);
            p.drawText(hr, Qt::AlignCenter, QString::number(r + 1));

            y += h;
            if (y > viewport()->height()) break;
        }
    }
}

void SpreadsheetView::drawSelection(QPainter& p)
{
    // Highlight selection range (light fill)
    if (!m_selection.isSingleCell()) {
        for (int r = m_selection.r1; r <= m_selection.r2; ++r) {
            for (int c = m_selection.c1; c <= m_selection.c2; ++c) {
                QRect cr = cellRect(r, c);
                if (cr.right() < HEADER_COL_W || cr.bottom() < HEADER_ROW_H) continue;
                p.fillRect(cr.adjusted(1, 1, 0, 0), m_selectionFill);
            }
        }
    }

    // Active cell border
    QRect cr = cellRect(m_curRow, m_curCol);
    cr = cr.adjusted(1, 1, 0, 0);
    if (cr.right() < HEADER_COL_W || cr.bottom() < HEADER_ROW_H) return;

    p.setPen(QPen(m_selectionColor, 2));
    p.setBrush(Qt::NoBrush);
    p.drawRect(cr);

    // Selection border
    if (!m_selection.isSingleCell()) {
        QRect selRect;
        for (int r = m_selection.r1; r <= m_selection.r2; ++r) {
            for (int c = m_selection.c1; c <= m_selection.c2; ++c) {
                QRect cr2 = cellRect(r, c);
                selRect = selRect.united(cr2);
            }
        }
        p.setPen(QPen(m_selectionColor, 2, Qt::DashLine));
        p.setBrush(Qt::NoBrush);
        p.drawRect(selRect.adjusted(1, 1, -1, -1));
    }
}

// ── Selection ─────────────────────────────────────────────────────────────────
void SpreadsheetView::setCurrentCell(int row, int col, bool extend)
{
    if (m_editing) commitEdit();

    m_curRow = std::clamp(row, 0, (int)m_totalRows - 1);
    m_curCol = std::clamp(col, 0, m_totalCols - 1);

    if (!extend) {
        m_selection = { m_curRow, m_curCol, m_curRow, m_curCol };
    } else {
        m_selection.r1 = std::min(m_selection.r1, m_curRow);
        m_selection.c1 = std::min(m_selection.c1, m_curCol);
        m_selection.r2 = std::max(m_selection.r2, m_curRow);
        m_selection.c2 = std::max(m_selection.c2, m_curCol);
    }

    scrollToCell(m_curRow, m_curCol);
    emit currentCellChanged(m_curRow, m_curCol);
    emit selectionChanged(m_selection);
    viewport()->update();
}

void SpreadsheetView::selectRange(int r1, int c1, int r2, int c2)
{
    m_selection = { std::min(r1,r2), std::min(c1,c2), std::max(r1,r2), std::max(c1,c2) };
    m_curRow = m_selection.r1;
    m_curCol = m_selection.c1;
    emit selectionChanged(m_selection);
    viewport()->update();
}

void SpreadsheetView::selectRow(int row)    { selectRange(row, 0, row, m_totalCols - 1); }
void SpreadsheetView::selectCol(int col)    { selectRange(0, col, (int)m_totalRows - 1, col); }
void SpreadsheetView::selectAll()           { selectRange(0, 0, (int)m_totalRows - 1, m_totalCols - 1); }

bool SpreadsheetView::navigateToAddress(const QString& address)
{
    // Parse "A1", "B12", "Sheet1!C3"
    QString addr = address;
    int bang = addr.indexOf('!');
    if (bang >= 0) addr = addr.mid(bang + 1);

    QRegularExpression re(R"(([A-Za-z]+)(\d+))");
    auto m = re.match(addr);
    if (!m.hasMatch()) return false;

    int col = colIndex(m.captured(1));
    int row = m.captured(2).toInt() - 1;
    if (col < 0 || row < 0) return false;

    setCurrentCell(row, col);
    return true;
}

// ── Scroll to cell ────────────────────────────────────────────────────────────
void SpreadsheetView::scrollToCell(int row, int col)
{
    // Vertical: adjust m_firstRow so row is visible
    if (row < m_firstRow) {
        m_firstRow = row;
        m_subpixelY = 0;
        verticalScrollBar()->setValue(rowToScrollValue(m_firstRow));
    } else {
        // Check if row is below visible area
        int y = HEADER_ROW_H;
        for (qint64 r = m_firstRow; r <= row; ++r)
            y += rowHeight((int)r);
        while (y > viewport()->height() && m_firstRow < row) {
            y -= rowHeight((int)m_firstRow);
            m_firstRow++;
        }
        verticalScrollBar()->setValue(rowToScrollValue(m_firstRow));
    }

    // Horizontal: pixel based
    int cellX = xOfCol(col);
    if (cellX < HEADER_COL_W) {
        // Scroll left
        int hval = horizontalScrollBar()->value();
        int delta = HEADER_COL_W - cellX;
        horizontalScrollBar()->setValue(std::max(0, hval - delta));
    } else if (cellX + colWidth(col) > viewport()->width()) {
        // Scroll right
        int hval = horizontalScrollBar()->value();
        int delta = cellX + colWidth(col) - viewport()->width();
        horizontalScrollBar()->setValue(hval + delta);
    }
}

// ── Resize event ──────────────────────────────────────────────────────────────
void SpreadsheetView::resizeEvent(QResizeEvent* e)
{
    QAbstractScrollArea::resizeEvent(e);
    updateScrollBars();
    if (m_editing) placeEditor();
}

// ── Key events ────────────────────────────────────────────────────────────────
void SpreadsheetView::keyPressEvent(QKeyEvent* e)
{
    bool ctrl  = e->modifiers() & Qt::ControlModifier;
    bool shift = e->modifiers() & Qt::ShiftModifier;

    if (m_editing) {
        // Escape cancels, Enter commits (handled by editor's returnPressed)
        if (e->key() == Qt::Key_Escape) { cancelEdit(); return; }
        if (e->key() == Qt::Key_Tab)    { commitEdit(); setCurrentCell(m_curRow, m_curCol + 1); return; }
        // Pass everything else to editor
        return;
    }

    auto move = [&](int dr, int dc) {
        setCurrentCell(m_curRow + dr, m_curCol + dc, shift);
    };

    switch (e->key()) {
    case Qt::Key_Up:        move(-1, 0); break;
    case Qt::Key_Down:      move( 1, 0); break;
    case Qt::Key_Left:      move(0, -1); break;
    case Qt::Key_Right:     move(0,  1); break;
    case Qt::Key_Tab:       move(0, shift ? -1 : 1); break;
    case Qt::Key_Return:
    case Qt::Key_Enter:     move(1, 0); break;
    case Qt::Key_PageUp:    move(-20, 0); break;
    case Qt::Key_PageDown:  move( 20, 0); break;
    case Qt::Key_Home:
        if (ctrl) setCurrentCell(0, 0);
        else      setCurrentCell(m_curRow, 0);
        break;
    case Qt::Key_End:
        if (ctrl) setCurrentCell(m_core ? m_core->rowCount(m_sheetId)-1 : 0,
                                 m_core ? m_core->columnCount(m_sheetId)-1 : 0);
        else      setCurrentCell(m_curRow, m_totalCols - 1);
        break;

    case Qt::Key_Delete:
    case Qt::Key_Backspace:
        deleteSelection();
        break;

    case Qt::Key_F2:
        startEditing();
        break;

    default:
        // Printable character → start editing
        if (!e->text().isEmpty() && !ctrl) {
            startEditing(e->text());
        }
        break;
    }
}

// ── Mouse events ──────────────────────────────────────────────────────────────
void SpreadsheetView::mousePressEvent(QMouseEvent* e)
{
    if (m_editing) commitEdit();

    if (e->button() == Qt::LeftButton) {
        // Check header resize
        int rc = colResizeAt(e->pos());
        int rr = rowResizeAt(e->pos());
        if (rc >= 0) {
            m_resizingCol   = true;
            m_resizeCol     = rc;
            m_resizeStart   = e->x();
            m_resizeOrigSize = colWidth(rc);
            viewport()->setCursor(Qt::SizeHorCursor);
            return;
        }
        if (rr >= 0) {
            m_resizingRow   = true;
            m_resizeRow     = rr;
            m_resizeStart   = e->y();
            m_resizeOrigSize = rowHeight(rr);
            viewport()->setCursor(Qt::SizeVerCursor);
            return;
        }

        // Click in column header → select column
        if (e->y() < HEADER_ROW_H && e->x() >= HEADER_COL_W) {
            int c = colAtX(e->x());
            if (c >= 0) selectCol(c);
            return;
        }
        // Click in row header → select row
        if (e->x() < HEADER_COL_W && e->y() >= HEADER_ROW_H) {
            int r = rowAtY(e->y());
            if (r >= 0) selectRow(r);
            return;
        }
        // Corner → select all
        if (e->x() < HEADER_COL_W && e->y() < HEADER_ROW_H) {
            selectAll();
            return;
        }

        int r = rowAtY(e->y());
        int c = colAtX(e->x());
        if (r >= 0 && c >= 0) {
            m_selecting    = true;
            m_selectStart  = e->pos();
            setCurrentCell(r, c, shift_held(e));
        }
    }
}

void SpreadsheetView::mouseMoveEvent(QMouseEvent* e)
{
    if (m_resizingCol) {
        int newW = std::max(20, m_resizeOrigSize + (e->x() - m_resizeStart));
        setColWidth(m_resizeCol, newW);
        return;
    }
    if (m_resizingRow) {
        int newH = std::max(8, m_resizeOrigSize + (e->y() - m_resizeStart));
        setRowHeight(m_resizeRow, newH);
        return;
    }

    if (m_selecting && (e->buttons() & Qt::LeftButton)) {
        int r = rowAtY(e->y());
        int c = colAtX(e->x());
        if (r >= 0 && c >= 0) {
            m_selection.r2 = r;
            m_selection.c2 = c;
            m_selection.r1 = std::min(m_selection.r1, r);
            m_selection.c1 = std::min(m_selection.c1, c);
            viewport()->update();
        }
    }

    // Cursor shape near resize handles
    if (colResizeAt(e->pos()) >= 0)
        viewport()->setCursor(Qt::SizeHorCursor);
    else if (rowResizeAt(e->pos()) >= 0)
        viewport()->setCursor(Qt::SizeVerCursor);
    else
        viewport()->setCursor(Qt::ArrowCursor);
}

void SpreadsheetView::mouseReleaseEvent(QMouseEvent*)
{
    m_selecting   = false;
    m_resizingCol = false;
    m_resizingRow = false;
    viewport()->setCursor(Qt::ArrowCursor);
}

void SpreadsheetView::mouseDoubleClickEvent(QMouseEvent* e)
{
    int r = rowAtY(e->y());
    int c = colAtX(e->x());
    if (r >= 0 && c >= 0) {
        setCurrentCell(r, c);
        startEditing();
    }
}

void SpreadsheetView::wheelEvent(QWheelEvent* e)
{
    if (e->modifiers() & Qt::ControlModifier) {
        int delta = e->angleDelta().y() > 0 ? 10 : -10;
        setZoom(m_zoom + delta);
        return;
    }
    QAbstractScrollArea::wheelEvent(e);
}

void SpreadsheetView::contextMenuEvent(QContextMenuEvent* e)
{
    int r = rowAtY(e->y());
    int c = colAtX(e->x());
    if (r >= 0 && c >= 0) {
        emit contextMenuRequested(e->globalPos(), r, c);
    }
}

// ── Inline editing ────────────────────────────────────────────────────────────
void SpreadsheetView::startEditing(const QString& initial)
{
    if (!m_core) return;
    m_editRow = m_curRow;
    m_editCol = m_curCol;

    Cell cell = m_core->getCell(m_sheetId, m_editRow, m_editCol);
    m_editOriginal = cell.formula.isEmpty()
                   ? cell.rawValue.toString()
                   : "=" + cell.formula;

    if (initial.isEmpty()) {
        m_editor->setText(m_editOriginal);
        m_editor->selectAll();
    } else {
        // Replace content with the typed character
        m_editor->clear();
        m_editor->setText(initial);
    }

    placeEditor();
    m_editor->show();
    m_editor->setFocus();
    m_editing = true;
}

void SpreadsheetView::placeEditor()
{
    QRect cr = cellRect(m_editRow, m_editCol);
    m_editor->setGeometry(cr.adjusted(1, 1, 0, 0));
}

void SpreadsheetView::commitEdit()
{
    if (!m_editing) return;
    m_editing = false;

    QString text = m_editor->text();
    m_editor->hide();

    if (!m_core) return;

    if (text != m_editOriginal) {
        if (text.startsWith('=')) {
            m_core->setCellFormula(m_sheetId, m_editRow, m_editCol, text.mid(1));
        } else {
            // Try numeric
            bool ok;
            double d = text.toDouble(&ok);
            if (ok)
                m_core->setCellValue(m_sheetId, m_editRow, m_editCol, d);
            else
                m_core->setCellValue(m_sheetId, m_editRow, m_editCol, text);
        }
        emit cellEdited(m_editRow, m_editCol, text);
    }
    viewport()->update();
}

void SpreadsheetView::cancelEdit()
{
    if (!m_editing) return;
    m_editing = false;
    m_editor->hide();
    viewport()->update();
}

// ── Clipboard ─────────────────────────────────────────────────────────────────
void SpreadsheetView::copy()
{
    if (!m_core) return;
    m_clipRange  = m_selection;
    m_clipIsCut  = false;
    m_hasClip    = true;

    // Build tab-delimited text for system clipboard
    QString text;
    for (int r = m_selection.r1; r <= m_selection.r2; ++r) {
        for (int c = m_selection.c1; c <= m_selection.c2; ++c) {
            if (c > m_selection.c1) text += '\t';
            Cell cell = m_core->getCell(m_sheetId, r, c);
            text += cell.cachedValue.isNull() ? cell.rawValue.toString()
                                              : cell.cachedValue.toString();
        }
        text += '\n';
    }
    QApplication::clipboard()->setText(text);
    viewport()->update();
}

void SpreadsheetView::cut()
{
    copy();
    m_clipIsCut = true;
}

void SpreadsheetView::paste()
{
    if (!m_core) return;

    const QClipboard* cb = QApplication::clipboard();
    QString text = cb->text();
    if (text.isEmpty()) return;

    QStringList rows = text.split('\n');
    for (int r = 0; r < rows.size(); ++r) {
        QStringList cols = rows[r].split('\t');
        for (int c = 0; c < cols.size(); ++c) {
            int tr = m_curRow + r;
            int tc = m_curCol + c;
            if (tr >= m_totalRows || tc >= m_totalCols) continue;
            const QString& val = cols[c];
            if (val.startsWith('='))
                m_core->setCellFormula(m_sheetId, tr, tc, val.mid(1));
            else {
                bool ok;
                double d = val.toDouble(&ok);
                if (ok)
                    m_core->setCellValue(m_sheetId, tr, tc, d);
                else
                    m_core->setCellValue(m_sheetId, tr, tc, val);
            }
        }
    }

    if (m_clipIsCut && m_hasClip) {
        for (int r = m_clipRange.r1; r <= m_clipRange.r2; ++r)
            for (int c = m_clipRange.c1; c <= m_clipRange.c2; ++c)
                m_core->clearCell(m_sheetId, r, c);
        m_hasClip = false;
    }

    viewport()->update();
}

void SpreadsheetView::deleteSelection()
{
    if (!m_core) return;
    for (int r = m_selection.r1; r <= m_selection.r2; ++r)
        for (int c = m_selection.c1; c <= m_selection.c2; ++c)
            m_core->clearCell(m_sheetId, r, c);
    viewport()->update();
}

// ── Sort ──────────────────────────────────────────────────────────────────────
void SpreadsheetView::sortAscending()
{
    if (!m_core) return;
    m_core->sortRange(m_sheetId,
                      m_selection.r1, m_selection.c1,
                      m_selection.r2, m_selection.c2,
                      m_selection.c1, Qt::AscendingOrder);
    refresh();
}

void SpreadsheetView::sortDescending()
{
    if (!m_core) return;
    m_core->sortRange(m_sheetId,
                      m_selection.r1, m_selection.c1,
                      m_selection.r2, m_selection.c2,
                      m_selection.c1, Qt::DescendingOrder);
    refresh();
}

// ── Find ──────────────────────────────────────────────────────────────────────
bool SpreadsheetView::findNext(const QString& text, bool caseSensitive, bool useRegex)
{
    if (!m_core || text.isEmpty()) return false;

    Qt::CaseSensitivity cs = caseSensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;

    int startRow = m_findRow;
    int startCol = m_findCol + 1;

    for (int r = startRow; r < m_core->rowCount(m_sheetId); ++r) {
        int colStart = (r == startRow) ? startCol : 0;
        for (int c = colStart; c < m_core->columnCount(m_sheetId); ++c) {
            Cell cell = m_core->getCell(m_sheetId, r, c);
            QString val = cell.cachedValue.isNull() ? cell.rawValue.toString()
                                                    : cell.cachedValue.toString();
            bool found = useRegex
                ? val.contains(QRegularExpression(text, caseSensitive
                    ? QRegularExpression::NoPatternOption
                    : QRegularExpression::CaseInsensitiveOption))
                : val.contains(text, cs);
            if (found) {
                m_findRow = r;
                m_findCol = c;
                setCurrentCell(r, c);
                return true;
            }
        }
    }
    m_findRow = 0; m_findCol = -1;
    return false;
}

void SpreadsheetView::replaceAll(const QString& find, const QString& replace,
                                 bool caseSensitive, bool useRegex)
{
    if (!m_core) return;
    Qt::CaseSensitivity cs = caseSensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;

    for (int r = 0; r < m_core->rowCount(m_sheetId); ++r) {
        for (int c = 0; c < m_core->columnCount(m_sheetId); ++c) {
            Cell cell = m_core->getCell(m_sheetId, r, c);
            QString val = cell.rawValue.toString();
            QString newVal;
            if (useRegex)
                newVal = val.replace(QRegularExpression(find, caseSensitive
                    ? QRegularExpression::NoPatternOption
                    : QRegularExpression::CaseInsensitiveOption), replace);
            else
                newVal = val.replace(find, replace, cs);
            if (newVal != val)
                m_core->setCellValue(m_sheetId, r, c, newVal);
        }
    }
    refresh();
}

// ── Refresh ───────────────────────────────────────────────────────────────────
void SpreadsheetView::refresh()
{
    viewport()->update();
}

void SpreadsheetView::refreshCell(int sheet, int row, int col)
{
    if (sheet != m_sheetId) return;
    QRect cr = cellRect(row, col);
    viewport()->update(cr);
}

// ── Header resize detection ───────────────────────────────────────────────────
int SpreadsheetView::colResizeAt(const QPoint& p) const
{
    if (p.y() > HEADER_ROW_H) return -1;
    if (p.x() < HEADER_COL_W) return -1;

    int x = HEADER_COL_W - m_subpixelX;
    for (int c = m_firstCol; c < m_totalCols; ++c) {
        x += colWidth(c);
        if (std::abs(p.x() - x) <= 4) return c;
        if (x > viewport()->width()) break;
    }
    return -1;
}

int SpreadsheetView::rowResizeAt(const QPoint& p) const
{
    if (p.x() > HEADER_COL_W) return -1;
    if (p.y() < HEADER_ROW_H) return -1;

    int y = HEADER_ROW_H - m_subpixelY;
    for (qint64 r = m_firstRow; r < m_totalRows; ++r) {
        y += rowHeight((int)r);
        if (std::abs(p.y() - y) <= 4) return (int)r;
        if (y > viewport()->height()) break;
    }
    return -1;
}

// ── Helpers ───────────────────────────────────────────────────────────────────
bool SpreadsheetView::shift_held(QMouseEvent* e)
{
    return e->modifiers() & Qt::ShiftModifier;
}
