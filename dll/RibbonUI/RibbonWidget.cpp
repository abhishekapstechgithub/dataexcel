#include "RibbonWidget.h"
#include "ColorButton.h"
#include <QTabWidget>
#include <QToolBar>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGroupBox>
#include <QComboBox>
#include <QToolButton>
#include <QLabel>
#include <QFontComboBox>
#include <QSpinBox>
#include <QFrame>
#include <QSizePolicy>

// ── Helper: create a styled tool button ──────────────────────────────────────
static QToolButton* makeBtn(const QString& icon, const QString& tip,
                            bool checkable = false)
{
    auto* btn = new QToolButton;
    btn->setIcon(QIcon::fromTheme(icon));
    // Fallback: text abbreviation when no icon
    if (btn->icon().isNull()) btn->setText(icon.section('-', -1).left(3));
    btn->setToolTip(tip);
    btn->setCheckable(checkable);
    btn->setAutoRaise(true);
    btn->setFixedSize(32, 32);
    return btn;
}

static QFrame* vSep() {
    auto* f = new QFrame;
    f->setFrameShape(QFrame::VLine);
    f->setFrameShadow(QFrame::Sunken);
    f->setFixedWidth(2);
    return f;
}

// ── RibbonWidget::Impl ────────────────────────────────────────────────────────
struct RibbonWidget::Impl {
    // Font controls
    QFontComboBox* fontFamily  { nullptr };
    QSpinBox*      fontSize    { nullptr };
    QToolButton*   btnBold     { nullptr };
    QToolButton*   btnItalic   { nullptr };
    QToolButton*   btnUnderline{ nullptr };
    ColorButton*   btnTextColor{ nullptr };
    ColorButton*   btnFillColor{ nullptr };
    // Alignment
    QToolButton*   btnAlignL   { nullptr };
    QToolButton*   btnAlignC   { nullptr };
    QToolButton*   btnAlignR   { nullptr };
    QToolButton*   btnVTop     { nullptr };
    QToolButton*   btnVMid     { nullptr };
    QToolButton*   btnVBot     { nullptr };
    QToolButton*   btnWrap     { nullptr };
    QToolButton*   btnMerge    { nullptr };
    // Number
    QComboBox*     numFormat   { nullptr };
    QToolButton*   btnDecInc   { nullptr };
    QToolButton*   btnDecDec   { nullptr };
};

// ── Construction ──────────────────────────────────────────────────────────────
RibbonWidget::RibbonWidget(QWidget* parent)
    : QWidget(parent), d(new Impl)
{
    setFixedHeight(110);
    setStyleSheet(R"(
        QWidget { background: #f3f2f1; }
        QToolButton { border: 1px solid transparent; border-radius: 3px;
                      background: transparent; font-size: 11px; }
        QToolButton:hover { background: #e5e4e3; border-color: #c8c6c4; }
        QToolButton:pressed, QToolButton:checked {
            background: #c7c5c3; border-color: #a09e9c; }
        QGroupBox { border: none; font-size: 10px; color: #555;
                    margin-top: 4px; padding-top: 2px; }
        QGroupBox::title { subcontrol-origin: margin; subcontrol-position: bottom center; }
        QComboBox { border: 1px solid #c8c6c4; border-radius: 2px;
                    background: white; padding: 1px 4px; min-width: 40px; }
        QFontComboBox { border: 1px solid #c8c6c4; border-radius: 2px;
                        background: white; min-width: 120px; }
        QSpinBox { border: 1px solid #c8c6c4; border-radius: 2px;
                   background: white; min-width: 44px; }
    )");

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(0,0,0,0);
    mainLayout->setSpacing(0);

    // ── Tab bar (Home only for now) ──────────────────────────────────────────
    auto* tabs = new QTabWidget;
    tabs->setTabPosition(QTabWidget::North);
    tabs->setDocumentMode(true);
    tabs->setStyleSheet("QTabWidget::pane { border:none; } "
                        "QTabBar::tab { padding:4px 12px; font-size:11px; } "
                        "QTabBar::tab:selected { background:#f3f2f1; border-bottom:2px solid #217346; }");

    // ── Home tab ─────────────────────────────────────────────────────────────
    auto* homeTab = new QWidget;
    auto* rowLayout = new QHBoxLayout(homeTab);
    rowLayout->setContentsMargins(4,2,4,2);
    rowLayout->setSpacing(0);

    // ── CLIPBOARD group ──────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Clipboard");
        auto* gl = new QHBoxLayout(grp);
        gl->setContentsMargins(4,4,4,16);
        gl->setSpacing(2);

        auto* btnPaste = makeBtn("edit-paste",  "Paste (Ctrl+V)");
        btnPaste->setFixedSize(46, 46);
        btnPaste->setText("Paste");
        btnPaste->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);

        auto* btnCut   = makeBtn("edit-cut",    "Cut (Ctrl+X)");
        auto* btnCopy  = makeBtn("edit-copy",   "Copy (Ctrl+C)");
        auto* btnFmtP  = makeBtn("format-indent-more", "Format Painter");
        btnCut->setText("Cut"); btnCopy->setText("Copy"); btnFmtP->setText("FmtP");

        auto* stack = new QVBoxLayout;
        stack->addWidget(btnCut);
        stack->addWidget(btnCopy);
        stack->addWidget(btnFmtP);

        gl->addWidget(btnPaste);
        gl->addLayout(stack);

        connect(btnPaste, &QToolButton::clicked, this, &RibbonWidget::pasteRequested);
        connect(btnCut,   &QToolButton::clicked, this, &RibbonWidget::cutRequested);
        connect(btnCopy,  &QToolButton::clicked, this, &RibbonWidget::copyRequested);
        connect(btnFmtP,  &QToolButton::clicked, this, &RibbonWidget::formatPainterRequested);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── FONT group ───────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Font");
        auto* gl = new QVBoxLayout(grp);
        gl->setContentsMargins(4,4,4,16);
        gl->setSpacing(2);

        // Row 1: family + size
        auto* row1 = new QHBoxLayout;
        d->fontFamily = new QFontComboBox;
        d->fontFamily->setCurrentFont(QFont("Calibri"));
        d->fontFamily->setToolTip("Font Family");

        d->fontSize = new QSpinBox;
        d->fontSize->setRange(6, 96);
        d->fontSize->setValue(11);
        d->fontSize->setToolTip("Font Size");

        row1->addWidget(d->fontFamily);
        row1->addWidget(d->fontSize);

        // Row 2: B / I / U / colors
        auto* row2 = new QHBoxLayout;
        d->btnBold      = makeBtn("format-text-bold",      "Bold (Ctrl+B)", true);
        d->btnItalic    = makeBtn("format-text-italic",    "Italic (Ctrl+I)", true);
        d->btnUnderline = makeBtn("format-text-underline", "Underline (Ctrl+U)", true);
        d->btnBold->setText("B");
        d->btnItalic->setText("I");
        d->btnUnderline->setText("U");

        d->btnTextColor = new ColorButton("A", grp);
        d->btnTextColor->setToolTip("Font Color");
        d->btnFillColor = new ColorButton("▥", grp);
        d->btnFillColor->setColor(Qt::yellow);
        d->btnFillColor->setToolTip("Fill Color");

        row2->addWidget(d->btnBold);
        row2->addWidget(d->btnItalic);
        row2->addWidget(d->btnUnderline);
        row2->addWidget(vSep());
        row2->addWidget(d->btnTextColor);
        row2->addWidget(d->btnFillColor);
        row2->addStretch();

        gl->addLayout(row1);
        gl->addLayout(row2);

        connect(d->fontFamily, &QFontComboBox::currentFontChanged, this,
                [this](const QFont& f){ emit fontFamilyChanged(f.family()); });
        connect(d->fontSize, qOverload<int>(&QSpinBox::valueChanged), this,
                &RibbonWidget::fontSizeChanged);
        connect(d->btnBold,      &QToolButton::toggled, this, &RibbonWidget::boldToggled);
        connect(d->btnItalic,    &QToolButton::toggled, this, &RibbonWidget::italicToggled);
        connect(d->btnUnderline, &QToolButton::toggled, this, &RibbonWidget::underlineToggled);
        connect(d->btnTextColor, &ColorButton::colorChanged, this, &RibbonWidget::textColorChanged);
        connect(d->btnFillColor, &ColorButton::colorChanged, this, &RibbonWidget::fillColorChanged);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── ALIGNMENT group ──────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Alignment");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,4,4,16);
        gl->setSpacing(2);

        auto* row1 = new QHBoxLayout; // vertical align
        d->btnVTop = makeBtn("format-justify-left",   "Align Top",    true); d->btnVTop->setText("⬆");
        d->btnVMid = makeBtn("format-justify-center", "Align Middle", true); d->btnVMid->setText("≡");
        d->btnVBot = makeBtn("format-justify-right",  "Align Bottom", true); d->btnVBot->setText("⬇");
        d->btnWrap = makeBtn("format-indent-more",    "Wrap Text",    true); d->btnWrap->setText("↵");
        d->btnMerge= makeBtn("object-group",          "Merge Cells"); d->btnMerge->setText("⊞");
        row1->addWidget(d->btnVTop); row1->addWidget(d->btnVMid); row1->addWidget(d->btnVBot);
        row1->addWidget(vSep());
        row1->addWidget(d->btnWrap); row1->addWidget(d->btnMerge);
        row1->addStretch();

        auto* row2 = new QHBoxLayout; // horizontal align
        d->btnAlignL = makeBtn("format-justify-left",   "Align Left",   true); d->btnAlignL->setText("◀");
        d->btnAlignC = makeBtn("format-justify-center", "Center",       true); d->btnAlignC->setText("☰");
        d->btnAlignR = makeBtn("format-justify-right",  "Align Right",  true); d->btnAlignR->setText("▶");
        row2->addWidget(d->btnAlignL); row2->addWidget(d->btnAlignC); row2->addWidget(d->btnAlignR);
        row2->addStretch();

        gl->addLayout(row1);
        gl->addLayout(row2);

        connect(d->btnAlignL, &QToolButton::clicked, this, [this]{ emit hAlignChanged(0); });
        connect(d->btnAlignC, &QToolButton::clicked, this, [this]{ emit hAlignChanged(1); });
        connect(d->btnAlignR, &QToolButton::clicked, this, [this]{ emit hAlignChanged(2); });
        connect(d->btnVTop,   &QToolButton::clicked, this, [this]{ emit vAlignChanged(0); });
        connect(d->btnVMid,   &QToolButton::clicked, this, [this]{ emit vAlignChanged(1); });
        connect(d->btnVBot,   &QToolButton::clicked, this, [this]{ emit vAlignChanged(2); });
        connect(d->btnWrap,   &QToolButton::toggled, this, &RibbonWidget::wrapTextToggled);
        connect(d->btnMerge,  &QToolButton::clicked, this, &RibbonWidget::mergeCellsRequested);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── NUMBER group ─────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Number");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,4,4,16);
        gl->setSpacing(2);

        d->numFormat = new QComboBox;
        d->numFormat->addItems({"General","Number","Currency","Percentage","Scientific"});
        d->numFormat->setToolTip("Number Format");

        auto* row2 = new QHBoxLayout;
        auto* btnCurr = makeBtn("", "Currency"); btnCurr->setText("$");
        auto* btnPct  = makeBtn("", "Percent");  btnPct->setText("%");
        d->btnDecInc  = makeBtn("", "Increase Decimal"); d->btnDecInc->setText(".0+");
        d->btnDecDec  = makeBtn("", "Decrease Decimal"); d->btnDecDec->setText(".0-");

        row2->addWidget(btnCurr); row2->addWidget(btnPct);
        row2->addWidget(vSep());
        row2->addWidget(d->btnDecInc); row2->addWidget(d->btnDecDec);
        row2->addStretch();

        gl->addWidget(d->numFormat);
        gl->addLayout(row2);

        connect(d->numFormat, qOverload<int>(&QComboBox::currentIndexChanged),
                this, &RibbonWidget::numberFormatChanged);
        connect(btnCurr, &QToolButton::clicked, this, [this]{ d->numFormat->setCurrentIndex(2);
            emit numberFormatChanged(2); });
        connect(btnPct,  &QToolButton::clicked, this, [this]{ d->numFormat->setCurrentIndex(3);
            emit numberFormatChanged(3); });
        connect(d->btnDecInc, &QToolButton::clicked, this, &RibbonWidget::increaseDecimalRequested);
        connect(d->btnDecDec, &QToolButton::clicked, this, &RibbonWidget::decreaseDecimalRequested);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── CELLS group ──────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Cells");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,4,4,16);
        gl->setSpacing(2);

        auto* btnInsRow = makeBtn("list-add",    "Insert Row");    btnInsRow->setText("+Row");
        auto* btnDelRow = makeBtn("list-remove", "Delete Row");    btnDelRow->setText("-Row");
        auto* btnInsCol = makeBtn("list-add",    "Insert Column"); btnInsCol->setText("+Col");
        auto* btnDelCol = makeBtn("list-remove", "Delete Column"); btnDelCol->setText("-Col");
        auto* btnFmtCel = makeBtn("document-properties","Format Cells"); btnFmtCel->setText("Fmt");

        auto* row1 = new QHBoxLayout;
        row1->addWidget(btnInsRow); row1->addWidget(btnDelRow); row1->addStretch();
        auto* row2 = new QHBoxLayout;
        row2->addWidget(btnInsCol); row2->addWidget(btnDelCol);
        row2->addWidget(btnFmtCel); row2->addStretch();

        gl->addLayout(row1);
        gl->addLayout(row2);

        connect(btnInsRow, &QToolButton::clicked, this, &RibbonWidget::insertRowRequested);
        connect(btnDelRow, &QToolButton::clicked, this, &RibbonWidget::deleteRowRequested);
        connect(btnInsCol, &QToolButton::clicked, this, &RibbonWidget::insertColumnRequested);
        connect(btnDelCol, &QToolButton::clicked, this, &RibbonWidget::deleteColumnRequested);
        connect(btnFmtCel, &QToolButton::clicked, this, &RibbonWidget::formatCellsRequested);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── EDITING group ────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Editing");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,4,4,16);
        gl->setSpacing(2);

        auto* btnSum    = makeBtn("", "AutoSum");        btnSum->setText("Σ");
        auto* btnSortA  = makeBtn("", "Sort A→Z");      btnSortA->setText("A↑Z");
        auto* btnSortD  = makeBtn("", "Sort Z→A");      btnSortD->setText("Z↓A");
        auto* btnFilter = makeBtn("", "Filter");         btnFilter->setText("⊟ Filter");
        auto* btnFind   = makeBtn("edit-find-replace",  "Find & Replace"); btnFind->setText("🔍 Find");

        auto* row1 = new QHBoxLayout;
        row1->addWidget(btnSum); row1->addWidget(btnSortA); row1->addWidget(btnSortD);
        row1->addStretch();
        auto* row2 = new QHBoxLayout;
        row2->addWidget(btnFilter); row2->addWidget(btnFind);
        row2->addStretch();

        gl->addLayout(row1);
        gl->addLayout(row2);

        connect(btnSum,    &QToolButton::clicked, this, &RibbonWidget::autoSumRequested);
        connect(btnSortA,  &QToolButton::clicked, this, &RibbonWidget::sortAscRequested);
        connect(btnSortD,  &QToolButton::clicked, this, &RibbonWidget::sortDescRequested);
        connect(btnFilter, &QToolButton::clicked, this, &RibbonWidget::filterRequested);
        connect(btnFind,   &QToolButton::clicked, this, &RibbonWidget::findReplaceRequested);

        rowLayout->addWidget(grp);
    }

    rowLayout->addStretch();
    tabs->addTab(homeTab, "Home");

    // ── Extra stub tabs ──────────────────────────────────────────────────────
    tabs->addTab(new QWidget, "Insert");
    tabs->addTab(new QWidget, "Page Layout");
    tabs->addTab(new QWidget, "Formulas");
    tabs->addTab(new QWidget, "Data");
    tabs->addTab(new QWidget, "Review");
    tabs->addTab(new QWidget, "View");

    mainLayout->addWidget(tabs);
}

RibbonWidget::~RibbonWidget() { delete d; }

void RibbonWidget::setFormatState(const RibbonFormatState& s) {
    QSignalBlocker b1(d->fontFamily), b2(d->fontSize),
                   b3(d->btnBold),    b4(d->btnItalic),
                   b5(d->btnUnderline),b6(d->numFormat);

    d->fontFamily->setCurrentFont(QFont(s.fontFamily));
    d->fontSize->setValue(s.fontSize);
    d->btnBold->setChecked(s.bold);
    d->btnItalic->setChecked(s.italic);
    d->btnUnderline->setChecked(s.underline);
    d->btnTextColor->setColor(s.textColor);
    d->btnFillColor->setColor(s.fillColor);
    d->numFormat->setCurrentIndex(s.numberFormat);
    d->btnWrap->setChecked(s.wrapText);
    d->btnAlignL->setChecked(s.hAlign == 0);
    d->btnAlignC->setChecked(s.hAlign == 1);
    d->btnAlignR->setChecked(s.hAlign == 2);
}

extern "C" RIBBON_API RibbonWidget* createRibbonWidget(QWidget* parent) {
    return new RibbonWidget(parent);
}
