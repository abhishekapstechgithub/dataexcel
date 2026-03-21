#pragma once
#include <QDialog>
#include "ISpreadsheetCore.h"
class QFontComboBox; class QSpinBox; class QComboBox; class QCheckBox;

class FormatCellsDialog : public QDialog {
    Q_OBJECT
public:
    FormatCellsDialog(const CellFormat& fmt, QWidget* parent = nullptr);
    CellFormat result() const;
private:
    QFontComboBox* m_font;
    QSpinBox*      m_size;
    QCheckBox      *m_bold, *m_italic, *m_underline, *m_wrap;
    QComboBox*     m_numFmt;
    QSpinBox*      m_decimals;
    CellFormat     m_format;
};
