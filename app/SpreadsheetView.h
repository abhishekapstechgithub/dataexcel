#pragma once
#include <QTableView>
#include <QHeaderView>
#include "SpreadsheetTableModel.h"
#include "ISpreadsheetCore.h"

class SpreadsheetView : public QTableView {
    Q_OBJECT
public:
    explicit SpreadsheetView(ISpreadsheetCore* core, SheetId sheet,
                             QWidget* parent = nullptr);
    ~SpreadsheetView() override;

    SpreadsheetTableModel* model() const;
    void switchSheet(SheetId sheet);

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

private slots:
    void onCurrentChanged(const QModelIndex& cur, const QModelIndex& prev);

private:
    ISpreadsheetCore*      m_core;
    SpreadsheetTableModel* m_model;

    void applyToSelection(std::function<void(int,int)> fn);
    void setupHeaders();
};
