#pragma once
#include <QTableView>
#include <QHeaderView>
#include "SpreadsheetTableModel.h"
#include "ISpreadsheetCore.h"

// ── SpreadsheetView ───────────────────────────────────────────────────────────
// Excel-style QTableView with:
//  - Zoom support (font scaling)
//  - Column/row resize with double-click auto-fit
//  - Freeze panes support
//  - Excel-style header highlighting (selected col/row turns green)
//  - Multi-cell selection with copy/paste
//  - Context menu
class SpreadsheetView : public QTableView {
    Q_OBJECT
public:
    explicit SpreadsheetView(ISpreadsheetCore* core, SheetId sheet,
                             QWidget* parent = nullptr);
    ~SpreadsheetView() override;

    SpreadsheetTableModel* model() const;
    void switchSheet(SheetId sheet);
    void setZoomFactor(qreal factor);

    // Current cell info
    int currentRow() const;
    int currentCol() const;
    CellFormat currentCellFormat() const;

    // Apply ribbon actions to current selection
    void applyFormat(const CellFormat& fmt);
    void applyHAlign(int a);
    void applyVAlign(int a);
    void applyNumberFormat(int fmt);
    void applyBold(bool on);
    void applyItalic(bool on);
    void applyUnderline(bool on);
    void applyWrapText(bool on);
    void applyTextColor(const QColor& c);
    void applyFillColor(const QColor& c);
    void applyFontFamily(const QString& f);
    void applyFontSize(int size);

    void insertRow();
    void deleteRow();
    void insertColumn();
    void deleteColumn();
    void sortAsc();
    void sortDesc();
    void autoSum();
    void mergeSelected();

signals:
    void selectionFormatChanged(const CellFormat& fmt, const QString& cellRef);

protected:
    void keyPressEvent(QKeyEvent* e) override;
    void contextMenuEvent(QContextMenuEvent* e) override;
    void mouseDoubleClickEvent(QMouseEvent* e) override;
    void wheelEvent(QWheelEvent* e) override;

private slots:
    void onCurrentChanged(const QModelIndex& cur, const QModelIndex& prev);

private:
    ISpreadsheetCore*      m_core;
    SpreadsheetTableModel* m_model;
    qreal                  m_zoomFactor { 1.0 };
    int                    m_baseRowHeight  { 24 };
    int                    m_baseColWidth   { 80 };

    void applyToSelection(std::function<void(int,int)> fn);
    void setupHeaders();
    void refreshSizes();
    QString selectionToRef() const;
};
