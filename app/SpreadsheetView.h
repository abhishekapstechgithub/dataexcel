#pragma once
#include <QTableView>
#include <QHeaderView>
#include <functional>
#include "SpreadsheetTableModel.h"
#include "ISpreadsheetCore.h"

class SpreadsheetView : public QTableView {
    Q_OBJECT
public:
    explicit SpreadsheetView(ISpreadsheetCore* core, SheetId sheet, QWidget* parent=nullptr);
    ~SpreadsheetView() override;

    SpreadsheetTableModel* model() const;
    void switchSheet(SheetId sheet);

    int   currentRow() const;
    int   currentCol() const;
    CellFormat currentCellFormat() const;
    qreal zoomFactor() const;
    void  setZoomFactor(qreal factor);

    // Format apply
    void applyFormat(const CellFormat& fmt);
    void applyBold(bool on);
    void applyItalic(bool on);
    void applyUnderline(bool on);
    void applyWrapText(bool on);
    void applyTextColor(const QColor& c);
    void applyFillColor(const QColor& c);
    void applyFontFamily(const QString& f);
    void applyFontSize(int sz);
    void applyHAlign(int a);
    void applyVAlign(int a);
    void applyNumberFormat(int fmt);

    // Actions
    void insertRow();
    void deleteRow();
    void insertColumn();
    void deleteColumn();
    void sortAsc();
    void sortDesc();
    void autoSum();
    void mergeSelected();

signals:
    void selectionFormatChanged(const CellFormat& fmt, const QString& ref);
    void zoomChanged(int percent);
    void boldToggled(bool on);

protected:
    void keyPressEvent(QKeyEvent* e) override;
    void contextMenuEvent(QContextMenuEvent* e) override;
    void wheelEvent(QWheelEvent* e) override;

private slots:
    void onCurrentChanged(const QModelIndex& cur, const QModelIndex& prev);

private:
    ISpreadsheetCore*      m_core;
    SpreadsheetTableModel* m_model;
    int    m_baseColWidth  { 80 };
    int    m_baseRowHeight { 22 };
    qreal  m_zoomFactor    { 1.0 };

    void setupHeaders();
    void applyToSelection(std::function<void(int,int)> fn);
    void doSort(Qt::SortOrder order);
};
