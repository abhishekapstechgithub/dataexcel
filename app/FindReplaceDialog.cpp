#include "FindReplaceDialog.h"
#include <QLineEdit>
#include <QCheckBox>
#include <QPushButton>
#include <QGridLayout>
#include <QLabel>

FindReplaceDialog::FindReplaceDialog(QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle("Find & Replace");
    setMinimumWidth(360);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* gl = new QGridLayout(this);
    gl->addWidget(new QLabel("Find:"),    0, 0);
    gl->addWidget(new QLabel("Replace:"), 1, 0);

    m_findEdit    = new QLineEdit(this);
    m_replaceEdit = new QLineEdit(this);
    m_matchCase   = new QCheckBox("Match case", this);

    gl->addWidget(m_findEdit,    0, 1);
    gl->addWidget(m_replaceEdit, 1, 1);
    gl->addWidget(m_matchCase,   2, 0, 1, 2);

    auto* btnFind    = new QPushButton("Find Next",    this);
    auto* btnAll     = new QPushButton("Replace All",  this);
    auto* btnClose   = new QPushButton("Close",        this);

    gl->addWidget(btnFind,  3, 0);
    gl->addWidget(btnAll,   3, 1);
    gl->addWidget(btnClose, 4, 0, 1, 2);

    connect(btnFind,  &QPushButton::clicked, this, &FindReplaceDialog::onFindNext);
    connect(btnAll,   &QPushButton::clicked, this, &FindReplaceDialog::onReplaceAll);
    connect(btnClose, &QPushButton::clicked, this, &QDialog::hide);
}

void FindReplaceDialog::onFindNext()
{
    emit findNextRequested(m_findEdit->text(), m_matchCase->isChecked());
}

void FindReplaceDialog::onReplaceAll()
{
    emit replaceAllRequested(m_findEdit->text(), m_replaceEdit->text(),
                             m_matchCase->isChecked());
}
