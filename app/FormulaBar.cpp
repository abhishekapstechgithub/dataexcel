// ═══════════════════════════════════════════════════════════════════════════════
//  FormulaBar.cpp — Excel-style formula bar implementation
// ═══════════════════════════════════════════════════════════════════════════════
#include "FormulaBar.h"
#include <QHBoxLayout>
#include <QLineEdit>
#include <QToolButton>
#include <QLabel>
#include <QIcon>
#include <QFont>
#include <QPainter>
#include <QPixmap>
#include <QKeyEvent>
#include <QFrame>
#include <QSizePolicy>

// ── Small helper: make an "fx" icon ──────────────────────────────────────────
static QIcon makeFxIcon()
{
    QPixmap pm(28, 20);
    pm.fill(Qt::transparent);
    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing);
    QFont f("Segoe UI", 9, QFont::Bold, true);
    p.setFont(f);
    p.setPen(QColor(0x217346));
    p.drawText(pm.rect(), Qt::AlignCenter, "fx");
    return QIcon(pm);
}

// ── FormulaLineEdit: captures Escape key ─────────────────────────────────────
// NOTE: FormulaLineEdit has Q_OBJECT and is defined in FormulaLineEdit.h
// to satisfy AUTOMOC requirements (Q_OBJECT classes must be in .h or
// explicitly moc'd). We define it here as a local class without Q_OBJECT
// and use an event filter instead.
class FormulaLineEdit : public QLineEdit
{
public:
    using QLineEdit::QLineEdit;

    std::function<void()> onEscape;

protected:
    void keyPressEvent(QKeyEvent* e) override
    {
        if (e->key() == Qt::Key_Escape) {
            if (onEscape) onEscape();
        } else {
            QLineEdit::keyPressEvent(e);
        }
    }
};

// ── FormulaBar ────────────────────────────────────────────────────────────────
FormulaBar::FormulaBar(QWidget* parent)
    : QWidget(parent)
{
    setFixedHeight(28);
    setStyleSheet(
        "FormulaBar {"
        "  background: #F2F2F2;"
        "  border-bottom: 1px solid #D0D0D0;"
        "}"
    );

    auto* layout = new QHBoxLayout(this);
    layout->setContentsMargins(4, 2, 4, 2);
    layout->setSpacing(2);

    // ── Name box (cell address) ───────────────────────────────────────────────
    m_nameBox = new QLineEdit(this);
    m_nameBox->setObjectName("nameBox");
    m_nameBox->setText("A1");
    m_nameBox->setFixedWidth(80);
    m_nameBox->setAlignment(Qt::AlignCenter);
    m_nameBox->setFont(QFont("Calibri", 10));
    m_nameBox->setStyleSheet(
        "QLineEdit {"
        "  border: 1px solid #C0C0C0;"
        "  border-radius: 2px;"
        "  background: white;"
        "  padding: 0 4px;"
        "}"
        "QLineEdit:focus {"
        "  border: 1px solid #217346;"
        "}"
    );
    connect(m_nameBox, &QLineEdit::returnPressed,
            this, &FormulaBar::onNameBoxReturnPressed);

    // ── Separator ─────────────────────────────────────────────────────────────
    auto* sep = new QFrame(this);
    sep->setFrameShape(QFrame::VLine);
    sep->setFrameShadow(QFrame::Sunken);
    sep->setFixedWidth(6);
    sep->setStyleSheet("color: #C0C0C0;");

    // ── fx button ─────────────────────────────────────────────────────────────
    m_fxButton = new QToolButton(this);
    m_fxButton->setIcon(makeFxIcon());
    m_fxButton->setIconSize(QSize(28, 18));
    m_fxButton->setFixedSize(32, 22);
    m_fxButton->setToolTip("Insert Function (Shift+F3)");
    m_fxButton->setStyleSheet(
        "QToolButton {"
        "  border: 1px solid transparent;"
        "  border-radius: 2px;"
        "  background: transparent;"
        "}"
        "QToolButton:hover {"
        "  border: 1px solid #C0C0C0;"
        "  background: #E0E0E0;"
        "}"
        "QToolButton:pressed {"
        "  background: #D0D0D0;"
        "}"
    );
    connect(m_fxButton, &QToolButton::clicked,
            this, &FormulaBar::insertFunctionRequested);

    // ── Formula / value edit ──────────────────────────────────────────────────
    auto* fle = new FormulaLineEdit(this);
    m_formulaEdit = fle;
    m_formulaEdit->setObjectName("formulaEdit");
    m_formulaEdit->setFont(QFont("Calibri", 10));
    m_formulaEdit->setPlaceholderText("");
    m_formulaEdit->setStyleSheet(
        "QLineEdit {"
        "  border: none;"
        "  background: white;"
        "  padding: 0 4px;"
        "}"
        "QLineEdit:focus {"
        "  border-bottom: 1px solid #217346;"
        "}"
    );
    m_formulaEdit->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Preferred);
    connect(m_formulaEdit, &QLineEdit::returnPressed,
            this, &FormulaBar::onFormulaEditReturnPressed);
    fle->onEscape = [this]{ onFormulaEditEscapePressed(); };

    layout->addWidget(m_nameBox);
    layout->addWidget(sep);
    layout->addWidget(m_fxButton);
    layout->addWidget(m_formulaEdit);
}

void FormulaBar::setCellAddress(const QString& address)
{
    m_nameBox->setText(address);
}

void FormulaBar::setFormulaText(const QString& text)
{
    m_formulaEdit->setText(text);
}

QString FormulaBar::formulaText() const
{
    return m_formulaEdit->text();
}

void FormulaBar::clear()
{
    m_nameBox->setText("A1");
    m_formulaEdit->clear();
}

void FormulaBar::beginEdit()
{
    m_formulaEdit->setFocus();
    m_formulaEdit->selectAll();
    m_editing = true;
}

void FormulaBar::endEdit()
{
    m_editing = false;
}

void FormulaBar::onNameBoxReturnPressed()
{
    QString addr = m_nameBox->text().trimmed().toUpper();
    if (!addr.isEmpty())
        emit nameBoxNavigate(addr);
}

void FormulaBar::onFormulaEditReturnPressed()
{
    emit formulaCommitted(m_formulaEdit->text());
    m_editing = false;
}

void FormulaBar::onFormulaEditEscapePressed()
{
    emit editCancelled();
    m_editing = false;
}

bool FormulaBar::eventFilter(QObject* obj, QEvent* event)
{
    return QWidget::eventFilter(obj, event);
}
