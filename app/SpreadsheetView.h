#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetView.h — High-performance virtual spreadsheet grid
//
//  Architecture:
//   • Custom QAbstractScrollArea; paints only visible cells with QPainter
//   • Virtual scrollbars: maps billions of rows to scrollbar int range
//   • Row/col headers as side panels
//   • Selection model: single cell, range, row, column
//   • Column/row resize by header drag
//   • Freeze panes
//   • Inline cell editing
// ═══════════════════════════════════════════════════════════════════════════════
#include <QAbstractScrollArea>
#include <QSet>
#include <QHash>
#include <QList>
#include <QPoint>
#include <QRect>
#include <QString>
#include <QColor>
#include <QFont>
#include <QModelIndex>

class ISpreadsheetCore;
class TileCache;
class QLineEdit;

// ── Cell selection ────────────────────────────────────────────────────────────
struct CellRange {
    int r1 { 0 }, c1 { 0 };
    int r2 { 0 }, c2 { 0 };

    bool isValid() const { return r1 <= r2 && c1 <= c2; }
    bool contains(int r, int c) const {
        return r >= r1 && r <= r2 && c >= c1 && c <= c2;
    }
    bool isSingleCell() const { return r1 == r2 && c1 == c2; }
};

// ═════════════════════════════════════════════════════════════════════════════
class SpreadsheetView : public QAbstractScrollArea
{
    Q_OBJECT

public:
    // Grid limits
    static constexpr qint64 MAX_ROWS = 1'048'576LL;
    static constexpr int    MAX_COLS = 16384;

    // Default cell dimensions
    static constexpr int DEF_COL_W = 96;
    static constexpr int DEF_ROW_H = 22;
    static constexpr int HEADER_COL_W = 50;   // row-number column width
    static constexpr int HEADER_ROW_H = 22;   // column-letter row height

    explicit SpreadsheetView(QWidget* parent = nullptr);
    ~SpreadsheetView() override;

    // ── Data source ──────────────────────────────────────────────────────────
    void setCore(ISpreadsheetCore* core);
    void setSheet(int sheetId);
    int  currentSheet() const { return m_sheetId; }

    // Override total row count for virtual scrolling (huge files)
    void setTotalRows(qint64 rows);
    void setTotalCols(int cols);

    // ── Selection ────────────────────────────────────────────────────────────
    CellRange selection() const { return m_selection; }
    int currentRow() const { return m_curRow; }
    int currentCol() const { return m_curCol; }

    void setCurrentCell(int row, int col, bool extendSelection = false);
    void selectRange(int r1, int c1, int r2, int c2);
    void selectRow(int row);
    void selectCol(int col);
    void selectAll();

    // Navigate to a cell address string like "B5" or "AA1000"
    bool navigateToAddress(const QString& address);

    // ── Column / Row sizes ───────────────────────────────────────────────────
    int  colWidth(int col) const;
    int  rowHeight(int row) const;
    void setColWidth(int col, int width);
    void setRowHeight(int row, int height);

    // ── Freeze panes ─────────────────────────────────────────────────────────
    void setFreezeRow(int row);   // 0 = no freeze
    void setFreezeCol(int col);   // 0 = no freeze

    // ── Zoom ─────────────────────────────────────────────────────────────────
    void setZoom(int percent);   // 50–400
    int  zoom() const { return m_zoom; }

    // ── Clipboard ────────────────────────────────────────────────────────────
    void copy();
    void cut();
    void paste();
    void deleteSelection();

    // ── Sort ─────────────────────────────────────────────────────────────────
    void sortAscending();
    void sortDescending();

    // ── Find ─────────────────────────────────────────────────────────────────
    bool findNext(const QString& text, bool caseSensitive, bool regex);
    void replaceAll(const QString& find, const QString& replace,
                    bool caseSensitive, bool regex);

    // ── Scroll helpers ───────────────────────────────────────────────────────
    void scrollToCell(int row, int col);

    // ── State ────────────────────────────────────────────────────────────────
    bool isEditing() const { return m_editing; }
    void commitEdit();
    void cancelEdit();

signals:
    void currentCellChanged(int row, int col);
    void selectionChanged(const CellRange& sel);
    void cellEdited(int row, int col, const QString& newValue);
    void columnResized(int col, int newWidth);
    void rowResized(int row, int newHeight);
    void contextMenuRequested(const QPoint& globalPos, int row, int col);

public slots:
    void refresh();          // repaint all visible cells
    void refreshCell(int sheet, int row, int col);

protected:
    // QAbstractScrollArea overrides
    void paintEvent(QPaintEvent* e) override;
    void resizeEvent(QResizeEvent* e) override;
    void scrollContentsBy(int dx, int dy) override;
    void keyPressEvent(QKeyEvent* e) override;
    void mousePressEvent(QMouseEvent* e) override;
    void mouseMoveEvent(QMouseEvent* e) override;
    void mouseReleaseEvent(QMouseEvent* e) override;
    void mouseDoubleClickEvent(QMouseEvent* e) override;
    void wheelEvent(QWheelEvent* e) override;
    void contextMenuEvent(QContextMenuEvent* e) override;

private:
    // ── Coordinate helpers ────────────────────────────────────────────────────
    // Pixel → logical cell (returns -1 if in header area)
    int  rowAtY(int y) const;
    int  colAtX(int x) const;
    // Logical → pixel (top-left of cell in viewport coords)
    int  yOfRow(int row) const;
    int  xOfCol(int col) const;
    // Cell rect in viewport coords
    QRect cellRect(int row, int col) const;

    // ── Visible range ─────────────────────────────────────────────────────────
    int  firstVisibleRow() const;
    int  lastVisibleRow()  const;
    int  firstVisibleCol() const;
    int  lastVisibleCol()  const;

    // ── Drawing ───────────────────────────────────────────────────────────────
    void drawHeaders(QPainter& p, const QRect& clip);
    void drawCells(QPainter& p, const QRect& clip);
    void drawCell(QPainter& p, int row, int col);
    void drawSelection(QPainter& p);
    void drawGridLines(QPainter& p, const QRect& clip);

    // ── Column label (A, B, ..., XFD) ────────────────────────────────────────
    static QString colLabel(int col);
    static int     colIndex(const QString& label);

    // ── Scrollbar ─────────────────────────────────────────────────────────────
    void updateScrollBars();
    // Virtual scroll: map scrollbar value → first visible row
    qint64 scrollValueToRow(int value) const;
    int    rowToScrollValue(qint64 row) const;

    // ── Inline editor ─────────────────────────────────────────────────────────
    void startEditing(const QString& initialText = {});
    void placeEditor();

    // ── Header resize ─────────────────────────────────────────────────────────
    int  colResizeAt(const QPoint& p) const; // col index if near resize line, else -1
    int  rowResizeAt(const QPoint& p) const;

    // ── Data ──────────────────────────────────────────────────────────────────
    ISpreadsheetCore* m_core    { nullptr };
    int               m_sheetId { 0 };
    qint64            m_totalRows { MAX_ROWS };
    int               m_totalCols { MAX_COLS };

    // ── Selection ─────────────────────────────────────────────────────────────
    CellRange  m_selection  { 0, 0, 0, 0 };
    int        m_curRow { 0 };
    int        m_curCol { 0 };
    bool       m_selecting { false };
    QPoint     m_selectStart;

    // ── Col/Row sizes ─────────────────────────────────────────────────────────
    QHash<int, int>   m_colWidths;
    QHash<int, int>   m_rowHeights;
    int               m_defaultColW { DEF_COL_W };
    int               m_defaultRowH { DEF_ROW_H };

    // ── Freeze panes ──────────────────────────────────────────────────────────
    int  m_freezeRow { 0 };
    int  m_freezeCol { 0 };

    // ── Scroll state ──────────────────────────────────────────────────────────
    qint64  m_firstRow  { 0 };   // logical first visible row
    int     m_firstCol  { 0 };   // logical first visible col
    int     m_subpixelY { 0 };   // sub-row pixel offset (for smooth scroll)
    int     m_subpixelX { 0 };

    // ── Zoom ──────────────────────────────────────────────────────────────────
    int    m_zoom { 100 };
    double m_scale { 1.0 };

    // ── Inline editor ─────────────────────────────────────────────────────────
    QLineEdit* m_editor   { nullptr };
    bool       m_editing  { false };
    int        m_editRow  { 0 };
    int        m_editCol  { 0 };
    QString    m_editOriginal;

    // ── Header resize ─────────────────────────────────────────────────────────
    int  m_resizeCol { -1 };
    int  m_resizeRow { -1 };
    int  m_resizeStart { 0 };
    int  m_resizeOrigSize { 0 };
    bool m_resizingCol { false };
    bool m_resizingRow { false };

    // ── Copy/cut state ────────────────────────────────────────────────────────
    CellRange  m_clipRange;
    bool       m_clipIsCut { false };
    bool       m_hasClip   { false };

    // ── Colors ────────────────────────────────────────────────────────────────
    QColor m_gridColor          { 0xD0D7DE };
    QColor m_headerBg           { 0xF2F2F2 };
    QColor m_headerText         { 0x444444 };
    QColor m_activeHeaderBg     { 0x217346 };
    QColor m_activeHeaderText   { Qt::white };
    QColor m_selectionColor     { 0x217346 };
    QColor m_selectionFill      { 0xE0F2E9 };

    // Find state
    int  m_findRow { 0 };
    int  m_findCol { 0 };

    static bool shift_held(QMouseEvent* e);
};
