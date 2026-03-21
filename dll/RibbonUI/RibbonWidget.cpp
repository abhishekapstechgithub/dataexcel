#include "RibbonWidget.h"
#include "ColorButton.h"
#include <QTabWidget>
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
#include <QPainter>
#include <QPixmap>
#include <QSvgRenderer>
#include <QByteArray>

// ── SVG icon factory — every icon is inline SVG, zero theme dependency ───────
static QIcon svgIcon(const QByteArray& svg, int sz = 20) {
    QSvgRenderer r(svg);
    QPixmap pm(sz, sz);
    pm.fill(Qt::transparent);
    QPainter p(&pm);
    r.render(&p);
    return QIcon(pm);
}

// All icons use currentColor=#444 so they work on light backgrounds
namespace Icons {
// Clipboard
static QIcon paste() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="2" width="6" height="4" rx="1"/><path d="M7 4H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2h-2"/><rect x="7" y="10" width="10" height="8" rx="1"/></svg>)"); }
static QIcon cut()   { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="6" cy="20" r="2"/><circle cx="18" cy="20" r="2"/><path d="M5.5 6.5 12 13l6.5-6.5"/><path d="M6 18 12 13l6 5"/></svg>)"); }
static QIcon copy()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>)"); }
static QIcon fmtpaint() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M18 4H6a2 2 0 0 0-2 2v4a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2z"/><path d="M12 12v8"/><path d="M9 18h6"/></svg>)"); }
// Font
static QIcon bold()      { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><text x="4" y="19" font-family="serif" font-size="18" font-weight="bold" fill="#444">B</text></svg>)"); }
static QIcon italic()    { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><text x="7" y="19" font-family="serif" font-size="18" font-style="italic" fill="#444">I</text></svg>)"); }
static QIcon underline() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round"><text x="5" y="17" font-family="serif" font-size="15" font-weight="bold" fill="#444" stroke="none">U</text><line x1="4" y1="22" x2="20" y2="22"/></svg>)"); }
static QIcon fontcolor() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><text x="3" y="16" font-family="serif" font-size="15" font-weight="bold" fill="#444">A</text><rect x="3" y="18" width="14" height="3" fill="#e84040"/></svg>)"); }
static QIcon fillcolor() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M16 2 4 14l2 2 12-12-2-2z"/><path d="m4 14 2 4h4l2-2"/><rect x="2" y="19" width="14" height="3" fill="#f5c518" stroke="none"/></svg>)"); }
// Alignment
static QIcon alignL()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="10" x2="15" y2="10"/><line x1="3" y1="14" x2="21" y2="14"/><line x1="3" y1="18" x2="15" y2="18"/></svg>)"); }
static QIcon alignC()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="6" y1="10" x2="18" y2="10"/><line x1="3" y1="14" x2="21" y2="14"/><line x1="6" y1="18" x2="18" y2="18"/></svg>)"); }
static QIcon alignR()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="9" y1="10" x2="21" y2="10"/><line x1="3" y1="14" x2="21" y2="14"/><line x1="9" y1="18" x2="21" y2="18"/></svg>)"); }
static QIcon vTop()    { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round"><line x1="3" y1="4" x2="21" y2="4"/><line x1="8" y1="8" x2="8" y2="20"/><line x1="16" y1="8" x2="16" y2="20"/><line x1="8" y1="14" x2="16" y2="14"/></svg>)"); }
static QIcon vMid()    { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round"><line x1="3" y1="12" x2="21" y2="12"/><line x1="8" y1="4" x2="8" y2="20"/><line x1="16" y1="4" x2="16" y2="20"/></svg>)"); }
static QIcon vBot()    { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round"><line x1="3" y1="20" x2="21" y2="20"/><line x1="8" y1="4" x2="8" y2="20"/><line x1="16" y1="4" x2="16" y2="20"/><line x1="8" y1="12" x2="16" y2="12"/></svg>)"); }
static QIcon wrap()    { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="6" x2="21" y2="6"/><path d="M3 12h13a3 3 0 0 1 0 6H3"/><polyline points="7 15 3 18 7 21"/></svg>)"); }
static QIcon merge()   { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="12" y1="3" x2="12" y2="12"/></svg>)"); }
// Number
static QIcon currency(){ return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 1 0 0 7h5a3.5 3.5 0 1 1 0 7H6"/></svg>)"); }
static QIcon percent() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><line x1="19" y1="5" x2="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/></svg>)"); }
static QIcon decInc()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><text x="1" y="16" font-size="11" fill="#444">.00</text><text x="13" y="12" font-size="9" fill="#2d7d46">+</text></svg>", 22)"); }
static QIcon decDec()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><text x="1" y="16" font-size="11" fill="#444">.0</text><text x="13" y="12" font-size="9" fill="#c0392b">-</text></svg>)", 22); }
// Cells
static QIcon insRow()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="7" rx="1"/><rect x="3" y="14" width="18" height="7" rx="1" opacity="0.4"/><line x1="12" y1="17" x2="12" y2="17" stroke="#2d7d46" stroke-width="2.5"/><line x1="9" y1="17" x2="15" y2="17" stroke="#2d7d46" stroke-width="2"/><line x1="12" y1="14.5" x2="12" y2="19.5" stroke="#2d7d46" stroke-width="2"/></svg>)"); }
static QIcon delRow()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="7" rx="1"/><rect x="3" y="14" width="18" height="7" rx="1" opacity="0.4"/><line x1="9" y1="17" x2="15" y2="17" stroke="#c0392b" stroke-width="2"/></svg>)"); }
static QIcon insCol()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="18" rx="1"/><rect x="14" y="3" width="7" height="18" rx="1" opacity="0.4"/><line x1="17" y1="9" x2="17" y2="15" stroke="#2d7d46" stroke-width="2"/><line x1="14.5" y1="12" x2="19.5" y2="12" stroke="#2d7d46" stroke-width="2"/></svg>)"); }
static QIcon delCol()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="18" rx="1"/><rect x="14" y="3" width="7" height="18" rx="1" opacity="0.4"/><line x1="14.5" y1="12" x2="19.5" y2="12" stroke="#c0392b" stroke-width="2"/></svg>)"); }
static QIcon fmtCell() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="3" x2="9" y2="9"/></svg>)"); }
// Editing
static QIcon autosum() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><text x="3" y="18" font-size="18" font-family="serif" fill="#444">Σ</text></svg>)"); }
static QIcon sortAsc() { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="6" x2="10" y2="6"/><line x1="3" y1="12" x2="14" y2="12"/><line x1="3" y1="18" x2="18" y2="18"/><polyline points="18 3 21 6 18 9"/><line x1="21" y1="6" x2="15" y2="6"/></svg>)"); }
static QIcon sortDesc(){ return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><line x1="3" y1="6" x2="18" y2="6"/><line x1="3" y1="12" x2="14" y2="12"/><line x1="3" y1="18" x2="10" y2="18"/><polyline points="18 21 21 18 18 15"/><line x1="21" y1="18" x2="15" y2="18"/></svg>)"); }
static QIcon filter()  { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>)"); }
static QIcon find()    { return svgIcon(R"(<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#444" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="10" cy="10" r="7"/><line x1="21" y1="21" x2="15" y2="15"/></svg>)"); }
}

// ── Helper: create a styled icon tool button ──────────────────────────────────
static QToolButton* makeBtn(const QIcon& icon, const QString& tip,
                             bool checkable = false, int w = 34, int h = 34)
{
    auto* btn = new QToolButton;
    btn->setIcon(icon);
    btn->setIconSize(QSize(20, 20));
    btn->setToolTip(tip);
    btn->setCheckable(checkable);
    btn->setAutoRaise(true);
    btn->setFixedSize(w, h);
    btn->setToolButtonStyle(Qt::ToolButtonIconOnly);
    return btn;
}

static QToolButton* makeLargeBtn(const QIcon& icon, const QString& label,
                                  const QString& tip)
{
    auto* btn = new QToolButton;
    btn->setIcon(icon);
    btn->setIconSize(QSize(28, 28));
    btn->setToolTip(tip);
    btn->setText(label);
    btn->setAutoRaise(true);
    btn->setFixedSize(54, 58);
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    return btn;
}

static QFrame* vSep() {
    auto* f = new QFrame;
    f->setFrameShape(QFrame::VLine);
    f->setFrameShadow(QFrame::Sunken);
    f->setFixedWidth(1);
    f->setFixedHeight(64);
    f->setStyleSheet("color: #d0d0d0;");
    return f;
}

// ── RibbonWidget::Impl ────────────────────────────────────────────────────────
struct RibbonWidget::Impl {
    QFontComboBox* fontFamily  { nullptr };
    QSpinBox*      fontSize    { nullptr };
    QToolButton*   btnBold     { nullptr };
    QToolButton*   btnItalic   { nullptr };
    QToolButton*   btnUnderline{ nullptr };
    ColorButton*   btnTextColor{ nullptr };
    ColorButton*   btnFillColor{ nullptr };
    QToolButton*   btnAlignL   { nullptr };
    QToolButton*   btnAlignC   { nullptr };
    QToolButton*   btnAlignR   { nullptr };
    QToolButton*   btnVTop     { nullptr };
    QToolButton*   btnVMid     { nullptr };
    QToolButton*   btnVBot     { nullptr };
    QToolButton*   btnWrap     { nullptr };
    QToolButton*   btnMerge    { nullptr };
    QComboBox*     numFormat   { nullptr };
    QToolButton*   btnDecInc   { nullptr };
    QToolButton*   btnDecDec   { nullptr };
};

// ── Construction ──────────────────────────────────────────────────────────────
RibbonWidget::RibbonWidget(QWidget* parent)
    : QWidget(parent), d(new Impl)
{
    setFixedHeight(116);
    setStyleSheet(R"(
        QWidget#ribbonRoot { background: #ffffff; }
        QTabWidget::pane   { border: none; background: #ffffff; }
        QTabBar::tab {
            padding: 5px 16px 4px;
            font-size: 12px;
            font-weight: 500;
            color: #444;
            background: transparent;
            border: none;
            border-bottom: 3px solid transparent;
        }
        QTabBar::tab:selected {
            color: #1a6b35;
            border-bottom: 3px solid #1a6b35;
        }
        QTabBar::tab:hover:!selected { color: #1a6b35; background: #f0f8f3; }
        QToolButton {
            border: 1px solid transparent;
            border-radius: 4px;
            background: transparent;
            font-size: 11px;
            color: #333;
        }
        QToolButton:hover  { background: #e8f4ed; border-color: #b8d9c4; }
        QToolButton:pressed, QToolButton:checked { background: #c8e6d0; border-color: #2d7d46; }
        QGroupBox {
            border: none;
            font-size: 10px;
            color: #888;
            margin-top: 2px;
            padding-top: 0px;
            padding-bottom: 2px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: bottom center;
            padding: 0 4px;
        }
        QComboBox {
            border: 1px solid #d0d0d0;
            border-radius: 3px;
            background: white;
            padding: 2px 6px;
            font-size: 12px;
            min-width: 40px;
        }
        QComboBox:hover { border-color: #2d7d46; }
        QFontComboBox {
            border: 1px solid #d0d0d0;
            border-radius: 3px;
            background: white;
            font-size: 12px;
            min-width: 130px;
        }
        QFontComboBox:hover { border-color: #2d7d46; }
        QSpinBox {
            border: 1px solid #d0d0d0;
            border-radius: 3px;
            background: white;
            font-size: 12px;
            min-width: 46px;
        }
        QSpinBox:hover { border-color: #2d7d46; }
    )");

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(0, 0, 0, 0);
    mainLayout->setSpacing(0);

    // ── Tab widget ────────────────────────────────────────────────────────────
    auto* tabs = new QTabWidget;
    tabs->setObjectName("ribbonRoot");
    tabs->setTabPosition(QTabWidget::North);
    tabs->setDocumentMode(true);

    // ── Home tab ──────────────────────────────────────────────────────────────
    auto* homeTab = new QWidget;
    homeTab->setStyleSheet("background:#ffffff;");
    auto* rowLayout = new QHBoxLayout(homeTab);
    rowLayout->setContentsMargins(6, 2, 6, 0);
    rowLayout->setSpacing(0);

    // ── CLIPBOARD ─────────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Clipboard");
        auto* gl  = new QHBoxLayout(grp);
        gl->setContentsMargins(4, 2, 4, 14);
        gl->setSpacing(2);

        auto* btnPaste = makeLargeBtn(Icons::paste(), "Paste", "Paste (Ctrl+V)");
        auto* vstack   = new QVBoxLayout;
        vstack->setSpacing(1);
        auto* btnCut  = makeBtn(Icons::cut(),      "Cut (Ctrl+X)");
        auto* btnCopy = makeBtn(Icons::copy(),     "Copy (Ctrl+C)");
        auto* btnFmtP = makeBtn(Icons::fmtpaint(), "Format Painter");
        vstack->addWidget(btnCut);
        vstack->addWidget(btnCopy);
        vstack->addWidget(btnFmtP);

        gl->addWidget(btnPaste);
        gl->addLayout(vstack);

        connect(btnPaste, &QToolButton::clicked, this, &RibbonWidget::pasteRequested);
        connect(btnCut,   &QToolButton::clicked, this, &RibbonWidget::cutRequested);
        connect(btnCopy,  &QToolButton::clicked, this, &RibbonWidget::copyRequested);
        connect(btnFmtP,  &QToolButton::clicked, this, &RibbonWidget::formatPainterRequested);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── FONT ──────────────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Font");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4, 2, 4, 14);
        gl->setSpacing(3);

        // Row 1: family + size
        auto* row1 = new QHBoxLayout;
        row1->setSpacing(3);
        d->fontFamily = new QFontComboBox;
        d->fontFamily->setCurrentFont(QFont("Calibri"));
        d->fontFamily->setToolTip("Font Family");
        d->fontFamily->setFixedHeight(24);

        d->fontSize = new QSpinBox;
        d->fontSize->setRange(6, 96);
        d->fontSize->setValue(11);
        d->fontSize->setToolTip("Font Size");
        d->fontSize->setFixedHeight(24);
        d->fontSize->setFixedWidth(50);

        row1->addWidget(d->fontFamily);
        row1->addWidget(d->fontSize);

        // Row 2: B I U | text color | fill color
        auto* row2 = new QHBoxLayout;
        row2->setSpacing(1);
        d->btnBold      = makeBtn(Icons::bold(),      "Bold (Ctrl+B)",      true);
        d->btnItalic    = makeBtn(Icons::italic(),    "Italic (Ctrl+I)",    true);
        d->btnUnderline = makeBtn(Icons::underline(), "Underline (Ctrl+U)", true);
        d->btnTextColor = new ColorButton("A", grp);
        d->btnTextColor->setToolTip("Font Color");
        d->btnTextColor->setFixedSize(34, 34);
        d->btnFillColor = new ColorButton("▥", grp);
        d->btnFillColor->setColor(Qt::yellow);
        d->btnFillColor->setToolTip("Fill Color");
        d->btnFillColor->setFixedSize(34, 34);

        row2->addWidget(d->btnBold);
        row2->addWidget(d->btnItalic);
        row2->addWidget(d->btnUnderline);
        row2->addSpacing(4);
        row2->addWidget(d->btnTextColor);
        row2->addWidget(d->btnFillColor);
        row2->addStretch();

        gl->addLayout(row1);
        gl->addLayout(row2);

        connect(d->fontFamily, &QFontComboBox::currentFontChanged, this,
                [this](const QFont& f){ emit fontFamilyChanged(f.family()); });
        connect(d->fontSize, qOverload<int>(&QSpinBox::valueChanged),
                this, &RibbonWidget::fontSizeChanged);
        connect(d->btnBold,      &QToolButton::toggled, this, &RibbonWidget::boldToggled);
        connect(d->btnItalic,    &QToolButton::toggled, this, &RibbonWidget::italicToggled);
        connect(d->btnUnderline, &QToolButton::toggled, this, &RibbonWidget::underlineToggled);
        connect(d->btnTextColor, &ColorButton::colorChanged, this, &RibbonWidget::textColorChanged);
        connect(d->btnFillColor, &ColorButton::colorChanged, this, &RibbonWidget::fillColorChanged);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── ALIGNMENT ────────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Alignment");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4, 2, 4, 14);
        gl->setSpacing(3);

        auto* row1 = new QHBoxLayout; row1->setSpacing(1);
        d->btnVTop = makeBtn(Icons::vTop(),  "Align Top",    true);
        d->btnVMid = makeBtn(Icons::vMid(),  "Align Middle", true);
        d->btnVBot = makeBtn(Icons::vBot(),  "Align Bottom", true);
        d->btnWrap = makeBtn(Icons::wrap(),  "Wrap Text",    true);
        d->btnMerge= makeBtn(Icons::merge(), "Merge Cells");
        row1->addWidget(d->btnVTop);
        row1->addWidget(d->btnVMid);
        row1->addWidget(d->btnVBot);
        row1->addSpacing(4);
        row1->addWidget(d->btnWrap);
        row1->addWidget(d->btnMerge);
        row1->addStretch();

        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        d->btnAlignL = makeBtn(Icons::alignL(), "Align Left",  true);
        d->btnAlignC = makeBtn(Icons::alignC(), "Center",      true);
        d->btnAlignR = makeBtn(Icons::alignR(), "Align Right", true);
        row2->addWidget(d->btnAlignL);
        row2->addWidget(d->btnAlignC);
        row2->addWidget(d->btnAlignR);
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

    // ── NUMBER ───────────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Number");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4, 2, 4, 14);
        gl->setSpacing(3);

        d->numFormat = new QComboBox;
        d->numFormat->addItems({"General","Number","Currency","Percentage","Scientific"});
        d->numFormat->setToolTip("Number Format");
        d->numFormat->setFixedHeight(24);

        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        auto* btnCurr = makeBtn(Icons::currency(), "Currency Style");
        auto* btnPct  = makeBtn(Icons::percent(),  "Percent Style");
        d->btnDecInc  = makeBtn(Icons::decInc(),   "Increase Decimal");
        d->btnDecDec  = makeBtn(Icons::decDec(),   "Decrease Decimal");
        row2->addWidget(btnCurr);
        row2->addWidget(btnPct);
        row2->addSpacing(4);
        row2->addWidget(d->btnDecInc);
        row2->addWidget(d->btnDecDec);
        row2->addStretch();

        gl->addWidget(d->numFormat);
        gl->addLayout(row2);

        connect(d->numFormat, qOverload<int>(&QComboBox::currentIndexChanged),
                this, &RibbonWidget::numberFormatChanged);
        connect(btnCurr, &QToolButton::clicked, this, [this]{
            d->numFormat->setCurrentIndex(2); emit numberFormatChanged(2); });
        connect(btnPct, &QToolButton::clicked, this, [this]{
            d->numFormat->setCurrentIndex(3); emit numberFormatChanged(3); });
        connect(d->btnDecInc, &QToolButton::clicked, this, &RibbonWidget::increaseDecimalRequested);
        connect(d->btnDecDec, &QToolButton::clicked, this, &RibbonWidget::decreaseDecimalRequested);

        rowLayout->addWidget(grp);
        rowLayout->addWidget(vSep());
    }

    // ── CELLS ────────────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Cells");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4, 2, 4, 14);
        gl->setSpacing(3);

        auto* btnInsRow = makeBtn(Icons::insRow(), "Insert Row");
        auto* btnDelRow = makeBtn(Icons::delRow(), "Delete Row");
        auto* btnInsCol = makeBtn(Icons::insCol(), "Insert Column");
        auto* btnDelCol = makeBtn(Icons::delCol(), "Delete Column");
        auto* btnFmtCel = makeBtn(Icons::fmtCell(),"Format Cells");

        auto* row1 = new QHBoxLayout; row1->setSpacing(1);
        row1->addWidget(btnInsRow);
        row1->addWidget(btnDelRow);
        row1->addWidget(btnFmtCel);
        row1->addStretch();
        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        row2->addWidget(btnInsCol);
        row2->addWidget(btnDelCol);
        row2->addStretch();

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

    // ── EDITING ──────────────────────────────────────────────────────────────
    {
        auto* grp = new QGroupBox("Editing");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4, 2, 4, 14);
        gl->setSpacing(3);

        auto* btnSum    = makeBtn(Icons::autosum(), "AutoSum");
        auto* btnSortA  = makeBtn(Icons::sortAsc(), "Sort A to Z");
        auto* btnSortD  = makeBtn(Icons::sortDesc(),"Sort Z to A");
        auto* btnFilter = makeBtn(Icons::filter(),  "Filter");
        auto* btnFind   = makeBtn(Icons::find(),    "Find & Replace (Ctrl+H)");

        auto* row1 = new QHBoxLayout; row1->setSpacing(1);
        row1->addWidget(btnSum);
        row1->addWidget(btnSortA);
        row1->addWidget(btnSortD);
        row1->addStretch();
        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        row2->addWidget(btnFilter);
        row2->addWidget(btnFind);
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
    tabs->addTab(homeTab,       "Home");
    tabs->addTab(new QWidget,   "Insert");
    tabs->addTab(new QWidget,   "Page Layout");
    tabs->addTab(new QWidget,   "Formulas");
    tabs->addTab(new QWidget,   "Data");
    tabs->addTab(new QWidget,   "Review");
    tabs->addTab(new QWidget,   "View");

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
#include "RibbonWidget.moc"
