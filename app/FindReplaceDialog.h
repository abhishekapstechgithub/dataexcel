#pragma once
#include <QDialog>
#include "ISpreadsheetCore.h"

class QLineEdit;
class QCheckBox;

class FindReplaceDialog : public QDialog {
    Q_OBJECT
public:
    explicit FindReplaceDialog(ISpreadsheetCore* core, SheetId sheet, QWidget* parent = nullptr);
private:
    ISpreadsheetCore* m_core;
    SheetId m_sheet;
    QLineEdit* m_findEdit;
    QLineEdit* m_replaceEdit;
    QCheckBox* m_matchCase;
    int m_lastRow{0}, m_lastCol{0};
    void findNext();
    void replaceOne();
    void replaceAll();
};
