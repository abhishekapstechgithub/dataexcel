#include "FindReplaceDialog.h"
#include <QLineEdit>
#include <QCheckBox>
#include <QPushButton>
#include <QGridLayout>
#include <QLabel>
#include <QMessageBox>

FindReplaceDialog::FindReplaceDialog(ISpreadsheetCore* core, SheetId sheet, QWidget* parent)
    : QDialog(parent), m_core(core), m_sheet(sheet)
{
    setWindowTitle("Find & Replace");
    setMinimumWidth(360);
    auto* gl = new QGridLayout(this);
    gl->addWidget(new QLabel("Find:"),    0,0);
    gl->addWidget(new QLabel("Replace:"), 1,0);
    m_findEdit    = new QLineEdit; gl->addWidget(m_findEdit,    0,1);
    m_replaceEdit = new QLineEdit; gl->addWidget(m_replaceEdit, 1,1);
    m_matchCase   = new QCheckBox("Match case"); gl->addWidget(m_matchCase, 2,0,1,2);

    auto* btnFind    = new QPushButton("Find Next");
    auto* btnReplace = new QPushButton("Replace");
    auto* btnAll     = new QPushButton("Replace All");
    auto* btnClose   = new QPushButton("Close");
    gl->addWidget(btnFind,    3,0); gl->addWidget(btnReplace, 3,1);
    gl->addWidget(btnAll,     4,0); gl->addWidget(btnClose,   4,1);

    connect(btnFind,    &QPushButton::clicked, this, &FindReplaceDialog::findNext);
    connect(btnReplace, &QPushButton::clicked, this, &FindReplaceDialog::replaceOne);
    connect(btnAll,     &QPushButton::clicked, this, &FindReplaceDialog::replaceAll);
    connect(btnClose,   &QPushButton::clicked, this, &QDialog::accept);
}

void FindReplaceDialog::findNext() {
    QString needle = m_findEdit->text();
    if (needle.isEmpty()) return;
    Qt::CaseSensitivity cs = m_matchCase->isChecked() ? Qt::CaseSensitive : Qt::CaseInsensitive;
    int rows = m_core->rowCount(m_sheet), cols = m_core->columnCount(m_sheet);
    for (int r = m_lastRow; r < rows; ++r) {
        int startCol = (r == m_lastRow) ? m_lastCol + 1 : 0;
        for (int c = startCol; c < cols; ++c) {
            Cell cell = m_core->getCell(m_sheet, r, c);
            if (cell.rawValue.toString().contains(needle, cs)) {
                m_lastRow = r; m_lastCol = c;
                // TODO: signal selection to main window
                return;
            }
        }
    }
    QMessageBox::information(this, "Find", "No more matches found.");
    m_lastRow = 0; m_lastCol = -1;
}

void FindReplaceDialog::replaceOne() {
    findNext();
    QString val = m_core->getCell(m_sheet, m_lastRow, m_lastCol).rawValue.toString();
    Qt::CaseSensitivity cs = m_matchCase->isChecked() ? Qt::CaseSensitive : Qt::CaseInsensitive;
    val.replace(m_findEdit->text(), m_replaceEdit->text(), cs);
    m_core->setCellValue(m_sheet, m_lastRow, m_lastCol, val);
}

void FindReplaceDialog::replaceAll() {
    QString needle = m_findEdit->text();
    if (needle.isEmpty()) return;
    Qt::CaseSensitivity cs = m_matchCase->isChecked() ? Qt::CaseSensitive : Qt::CaseInsensitive;
    int count = 0, rows = m_core->rowCount(m_sheet), cols = m_core->columnCount(m_sheet);
    for (int r = 0; r < rows; ++r)
        for (int c = 0; c < cols; ++c) {
            Cell cell = m_core->getCell(m_sheet, r, c);
            QString v = cell.rawValue.toString();
            if (v.contains(needle, cs)) {
                v.replace(needle, m_replaceEdit->text(), cs);
                m_core->setCellValue(m_sheet, r, c, v);
                count++;
            }
        }
    QMessageBox::information(this, "Replace All",
        QString("Replaced %1 occurrence(s).").arg(count));
}
