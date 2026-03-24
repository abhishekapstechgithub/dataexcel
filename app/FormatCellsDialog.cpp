#include "FormatCellsDialog.h"
// ISpreadsheetCore included via FormatCellsDialog.h
#include <QFontComboBox>
#include <QSpinBox>
#include <QComboBox>
#include <QCheckBox>
#include <QPushButton>
#include <QFormLayout>
#include <QVBoxLayout>
#include <QDialogButtonBox>
#include <QTabWidget>

FormatCellsDialog::FormatCellsDialog(const CellFormat& fmt, QWidget* parent)
    : QDialog(parent), m_format(fmt)
{
    setWindowTitle("Format Cells");
    setMinimumWidth(380);
    auto* tabs = new QTabWidget;

    // ── Font tab ──────────────────────────────────────────────────────────
    auto* fontTab = new QWidget;
    auto* fl = new QFormLayout(fontTab);
    m_font = new QFontComboBox; m_font->setCurrentFont(fmt.font);
    m_size = new QSpinBox; m_size->setRange(6,96); m_size->setValue(fmt.font.pointSize() > 0 ? fmt.font.pointSize() : 11);
    m_bold      = new QCheckBox("Bold");      m_bold->setChecked(fmt.bold);
    m_italic    = new QCheckBox("Italic");    m_italic->setChecked(fmt.italic);
    m_underline = new QCheckBox("Underline"); m_underline->setChecked(fmt.underline);
    fl->addRow("Font:",      m_font);
    fl->addRow("Size:",      m_size);
    fl->addRow(m_bold);
    fl->addRow(m_italic);
    fl->addRow(m_underline);
    tabs->addTab(fontTab, "Font");

    // ── Number tab ────────────────────────────────────────────────────────
    auto* numTab = new QWidget;
    auto* nl = new QFormLayout(numTab);
    m_numFmt  = new QComboBox;
    m_numFmt->addItems({"General","Number","Currency","Percentage","Scientific"});
    m_numFmt->setCurrentIndex(fmt.numberFormat);
    m_decimals = new QSpinBox; m_decimals->setRange(0,10); m_decimals->setValue(fmt.decimals);
    m_wrap = new QCheckBox("Wrap text"); m_wrap->setChecked(fmt.wrapText);
    nl->addRow("Format:",   m_numFmt);
    nl->addRow("Decimals:", m_decimals);
    nl->addRow(m_wrap);
    tabs->addTab(numTab, "Number");

    auto* buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel);
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);

    auto* vl = new QVBoxLayout(this);
    vl->addWidget(tabs);
    vl->addWidget(buttons);
}

CellFormat FormatCellsDialog::cellFormat() const {
    CellFormat fmt = m_format;
    fmt.font        = m_font->currentFont();
    fmt.font.setPointSize(m_size->value());
    fmt.bold        = m_bold->isChecked();
    fmt.italic      = m_italic->isChecked();
    fmt.underline   = m_underline->isChecked();
    fmt.numberFormat= m_numFmt->currentIndex();
    fmt.decimals    = m_decimals->value();
    fmt.wrapText    = m_wrap->isChecked();
    return fmt;
}
