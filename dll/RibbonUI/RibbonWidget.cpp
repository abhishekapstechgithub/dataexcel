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
#include <QPolygon>
#include <functional>

// ── Pure QPainter icon factory — no SVG module, no theme, no file I/O ────────
static QIcon makeIcon(std::function<void(QPainter&)> draw) {
    const int sz = 20;
    QPixmap pm(sz, sz);
    pm.fill(Qt::transparent);
    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing);
    draw(p);
    return QIcon(pm);
}

static QPen pen(const QColor& c = QColor("#444444"), qreal w = 1.6) {
    return QPen(c, w, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
}

namespace Icons {

static QIcon paste() { return makeIcon([](QPainter& p){
    p.setPen(pen()); p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(3,5,14,13,1,1);
    p.setBrush(QColor("#ffffff")); p.drawRoundedRect(7,2,6,4,1,1);
    p.setBrush(Qt::NoBrush);      p.drawRoundedRect(7,2,6,4,1,1);
    p.drawLine(6,10,14,10); p.drawLine(6,13,14,13);
}); }

static QIcon cut() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawEllipse(2,13,5,5); p.drawEllipse(13,13,5,5);
    p.drawLine(4,13,10,8); p.drawLine(16,13,10,8);
    p.drawLine(4,4,10,8);  p.drawLine(16,4,10,8);
}); }

static QIcon copy() { return makeIcon([](QPainter& p){
    p.setPen(pen()); p.setBrush(QColor("#f5f5f5"));
    p.drawRoundedRect(6,6,12,12,1,1);
    p.setBrush(QColor("#ffffff"));
    p.drawRoundedRect(2,2,12,12,1,1);
    p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(2,2,12,12,1,1);
}); }

static QIcon fmtpaint() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(3,4,17,4); p.drawLine(3,8,13,8);
    p.drawLine(3,12,17,12); p.drawLine(3,16,13,16);
    p.setPen(pen(QColor("#2d7d46"), 2.0));
    p.drawLine(13,13,18,18);
    p.drawEllipse(11,11,3,3);
}); }

static QIcon bold() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("serif"); f.setPixelSize(14); f.setBold(true);
    p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,20,20), Qt::AlignCenter, "B");
}); }

static QIcon italic() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("serif"); f.setPixelSize(14); f.setItalic(true);
    p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,20,20), Qt::AlignCenter, "I");
}); }

static QIcon underline() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("sans-serif"); f.setPixelSize(12); f.setBold(true);
    p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,-2,20,20), Qt::AlignCenter, "U");
    p.setPen(pen(QColor("#222"), 2.0));
    p.drawLine(3,18,17,18);
}); }

static QIcon fontcolor() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("sans-serif"); f.setPixelSize(12); f.setBold(true);
    p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,-2,20,20), Qt::AlignCenter, "A");
    p.setBrush(QColor("#e84040")); p.setPen(Qt::NoPen);
    p.drawRect(2,16,16,3);
}); }

static QIcon fillcolor() { return makeIcon([](QPainter& p){
    p.setPen(pen()); p.setBrush(QColor("#444"));
    QPolygon poly; poly<<QPoint(5,13)<<QPoint(10,3)<<QPoint(15,13)<<QPoint(12,13)<<QPoint(12,15)<<QPoint(8,15)<<QPoint(8,13);
    p.drawPolygon(poly);
    p.setBrush(QColor("#f5c518")); p.setPen(Qt::NoPen);
    p.drawRect(2,16,16,3);
}); }

static QIcon alignL() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,5,18,5); p.drawLine(2,9,13,9);
    p.drawLine(2,13,18,13); p.drawLine(2,17,13,17);
}); }

static QIcon alignC() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,5,18,5); p.drawLine(5,9,15,9);
    p.drawLine(2,13,18,13); p.drawLine(5,17,15,17);
}); }

static QIcon alignR() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,5,18,5); p.drawLine(7,9,18,9);
    p.drawLine(2,13,18,13); p.drawLine(7,17,18,17);
}); }

static QIcon vTop() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,3,18,3);
    p.drawLine(7,5,7,17); p.drawLine(13,5,13,17); p.drawLine(7,11,13,11);
}); }

static QIcon vMid() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,10,18,10);
    p.drawLine(7,3,7,17); p.drawLine(13,3,13,17);
}); }

static QIcon vBot() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,17,18,17);
    p.drawLine(7,3,7,15); p.drawLine(13,3,13,15); p.drawLine(7,9,13,9);
}); }

static QIcon wrap() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,5,18,5);
    p.drawLine(2,10,14,10);
    p.drawArc(11,8,7,7,0,180*16);
    p.drawLine(2,15,11,15);
    QPolygon arr; arr<<QPoint(5,11)<<QPoint(2,14)<<QPoint(5,17);
    p.drawPolyline(arr);
}); }

static QIcon merge() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawRect(2,2,16,16);
    p.drawLine(2,10,18,10); p.drawLine(10,2,10,10);
    p.setPen(pen(QColor("#2d7d46"), 1.8));
    QPolygon a; a<<QPoint(4,14)<<QPoint(7,12)<<QPoint(7,16); p.drawPolygon(a);
    QPolygon b; b<<QPoint(16,14)<<QPoint(13,12)<<QPoint(13,16); p.drawPolygon(b);
    p.drawLine(4,14,16,14);
}); }

static QIcon currency() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("sans-serif"); f.setPixelSize(13);
    p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,20,20), Qt::AlignCenter, "$");
}); }

static QIcon percent() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("sans-serif"); f.setPixelSize(12);
    p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,20,20), Qt::AlignCenter, "%");
}); }

static QIcon decInc() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("sans-serif"); f.setPixelSize(9);
    p.setFont(f); p.setPen(QColor("#444444")); p.drawText(1,14,".00");
    p.setPen(QColor("#2d7d46")); p.drawText(13,10,"+");
}); }

static QIcon decDec() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("sans-serif"); f.setPixelSize(9);
    p.setFont(f); p.setPen(QColor("#444444")); p.drawText(1,14,".0");
    p.setPen(QColor("#c0392b")); p.drawText(12,10,QString(QChar(0x2212)));
}); }

static QIcon insRow() { return makeIcon([](QPainter& p){
    p.setPen(pen()); p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(2,2,16,6,1,1);
    p.drawRoundedRect(2,12,16,6,1,1);
    p.setPen(pen(QColor("#2d7d46"), 2.0));
    p.drawLine(10,10,10,20); p.drawLine(7,15,13,15);
}); }

static QIcon delRow() { return makeIcon([](QPainter& p){
    p.setPen(pen()); p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(2,2,16,6,1,1);
    p.drawRoundedRect(2,12,16,6,1,1);
    p.setPen(pen(QColor("#c0392b"), 2.0));
    p.drawLine(7,15,13,15);
}); }

static QIcon insCol() { return makeIcon([](QPainter& p){
    p.setPen(pen()); p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(2,2,6,16,1,1);
    p.drawRoundedRect(12,2,6,16,1,1);
    p.setPen(pen(QColor("#2d7d46"), 2.0));
    p.drawLine(8,5,8,15); p.drawLine(5,10,11,10);
}); }

static QIcon delCol() { return makeIcon([](QPainter& p){
    p.setPen(pen()); p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(2,2,6,16,1,1);
    p.drawRoundedRect(12,2,6,16,1,1);
    p.setPen(pen(QColor("#c0392b"), 2.0));
    p.drawLine(5,10,11,10);
}); }

static QIcon fmtCell() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawRoundedRect(2,2,16,16,2,2);
    p.drawLine(2,8,18,8); p.drawLine(8,2,8,8);
}); }

static QIcon autosum() { return makeIcon([](QPainter& p){
    QFont f; f.setFamily("serif"); f.setPixelSize(15);
    p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,20,20), Qt::AlignCenter, QString(QChar(0x03A3)));
}); }

static QIcon sortAsc() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,5,10,5); p.drawLine(2,10,14,10); p.drawLine(2,15,18,15);
    p.setPen(pen(QColor("#2d7d46"), 1.8));
    p.drawLine(17,4,17,12);
    QPolygon arr; arr<<QPoint(14,7)<<QPoint(17,4)<<QPoint(20,7);
    p.drawPolyline(arr);
}); }

static QIcon sortDesc() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawLine(2,5,18,5); p.drawLine(2,10,14,10); p.drawLine(2,15,10,15);
    p.setPen(pen(QColor("#2d7d46"), 1.8));
    p.drawLine(17,4,17,12);
    QPolygon arr; arr<<QPoint(14,9)<<QPoint(17,12)<<QPoint(20,9);
    p.drawPolyline(arr);
}); }

static QIcon filter() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    QPolygon poly;
    poly<<QPoint(2,3)<<QPoint(18,3)<<QPoint(12,10)<<QPoint(12,17)<<QPoint(8,17)<<QPoint(8,10)<<QPoint(2,3);
    p.drawPolyline(poly);
}); }

static QIcon find() { return makeIcon([](QPainter& p){
    p.setPen(pen());
    p.drawEllipse(2,2,12,12);
    p.setPen(pen(QColor("#444"), 2.2));
    p.drawLine(12,12,18,18);
}); }

} // namespace Icons

// ── Helper: icon-only tool button ─────────────────────────────────────────────
static QToolButton* makeBtn(const QIcon& icon, const QString& tip,
                             bool checkable = false)
{
    auto* btn = new QToolButton;
    btn->setIcon(icon);
    btn->setIconSize(QSize(18, 18));
    btn->setToolTip(tip);
    btn->setCheckable(checkable);
    btn->setAutoRaise(true);
    btn->setFixedSize(30, 30);
    btn->setToolButtonStyle(Qt::ToolButtonIconOnly);
    return btn;
}

static QToolButton* makeLargeBtn(const QIcon& icon, const QString& label,
                                  const QString& tip)
{
    auto* btn = new QToolButton;
    btn->setIcon(icon);
    btn->setIconSize(QSize(26, 26));
    btn->setToolTip(tip);
    btn->setText(label);
    btn->setAutoRaise(true);
    btn->setFixedSize(50, 58);
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
        QWidget { background: #ffffff; }
        QTabWidget::pane { border: none; background: #ffffff; }
        QTabBar::tab {
            padding: 5px 16px 4px;
            font-size: 12px;
            font-weight: 500;
            color: #444;
            background: transparent;
            border: none;
            border-bottom: 3px solid transparent;
        }
        QTabBar::tab:selected { color: #1a6b35; border-bottom: 3px solid #1a6b35; }
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
            border: none; font-size: 10px; color: #888;
            margin-top: 2px; padding-top: 0px; padding-bottom: 2px;
        }
        QGroupBox::title { subcontrol-origin: margin; subcontrol-position: bottom center; padding: 0 4px; }
        QComboBox {
            border: 1px solid #d0d0d0; border-radius: 3px;
            background: white; padding: 2px 6px; font-size: 12px; min-width: 40px;
        }
        QComboBox:hover { border-color: #2d7d46; }
        QFontComboBox {
            border: 1px solid #d0d0d0; border-radius: 3px;
            background: white; font-size: 12px; min-width: 130px;
        }
        QFontComboBox:hover { border-color: #2d7d46; }
        QSpinBox {
            border: 1px solid #d0d0d0; border-radius: 3px;
            background: white; font-size: 12px; min-width: 46px;
        }
        QSpinBox:hover { border-color: #2d7d46; }
    )");

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(0, 0, 0, 0);
    mainLayout->setSpacing(0);

    auto* tabs = new QTabWidget;
    tabs->setTabPosition(QTabWidget::North);
    tabs->setDocumentMode(true);

    // ── Home tab ──────────────────────────────────────────────────────────────
    auto* homeTab = new QWidget;
    homeTab->setStyleSheet("background:#ffffff;");
    auto* rowLayout = new QHBoxLayout(homeTab);
    rowLayout->setContentsMargins(6, 2, 6, 0);
    rowLayout->setSpacing(0);

    // CLIPBOARD
    {
        auto* grp = new QGroupBox("Clipboard");
        auto* gl  = new QHBoxLayout(grp);
        gl->setContentsMargins(4,2,4,14); gl->setSpacing(2);
        auto* btnPaste = makeLargeBtn(Icons::paste(), "Paste", "Paste (Ctrl+V)");
        auto* vstack   = new QVBoxLayout; vstack->setSpacing(1);
        auto* btnCut  = makeBtn(Icons::cut(),      "Cut (Ctrl+X)");
        auto* btnCopy = makeBtn(Icons::copy(),     "Copy (Ctrl+C)");
        auto* btnFmtP = makeBtn(Icons::fmtpaint(), "Format Painter");
        vstack->addWidget(btnCut); vstack->addWidget(btnCopy); vstack->addWidget(btnFmtP);
        gl->addWidget(btnPaste); gl->addLayout(vstack);
        connect(btnPaste, &QToolButton::clicked, this, &RibbonWidget::pasteRequested);
        connect(btnCut,   &QToolButton::clicked, this, &RibbonWidget::cutRequested);
        connect(btnCopy,  &QToolButton::clicked, this, &RibbonWidget::copyRequested);
        connect(btnFmtP,  &QToolButton::clicked, this, &RibbonWidget::formatPainterRequested);
        rowLayout->addWidget(grp); rowLayout->addWidget(vSep());
    }

    // FONT
    {
        auto* grp = new QGroupBox("Font");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,2,4,14); gl->setSpacing(3);
        auto* row1 = new QHBoxLayout; row1->setSpacing(3);
        d->fontFamily = new QFontComboBox;
        d->fontFamily->setCurrentFont(QFont("Calibri"));
        d->fontFamily->setFixedHeight(24);
        d->fontSize = new QSpinBox;
        d->fontSize->setRange(6,96); d->fontSize->setValue(11);
        d->fontSize->setFixedHeight(24); d->fontSize->setFixedWidth(50);
        row1->addWidget(d->fontFamily); row1->addWidget(d->fontSize);
        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        d->btnBold      = makeBtn(Icons::bold(),      "Bold (Ctrl+B)",      true);
        d->btnItalic    = makeBtn(Icons::italic(),    "Italic (Ctrl+I)",    true);
        d->btnUnderline = makeBtn(Icons::underline(), "Underline (Ctrl+U)", true);
        d->btnTextColor = new ColorButton("A", grp); d->btnTextColor->setFixedSize(30,30);
        d->btnTextColor->setToolTip("Font Color");
        d->btnFillColor = new ColorButton(QString(QChar(0x25A6)), grp);
        d->btnFillColor->setColor(Qt::yellow); d->btnFillColor->setFixedSize(30,30);
        d->btnFillColor->setToolTip("Fill Color");
        row2->addWidget(d->btnBold); row2->addWidget(d->btnItalic);
        row2->addWidget(d->btnUnderline); row2->addSpacing(3);
        row2->addWidget(d->btnTextColor); row2->addWidget(d->btnFillColor);
        row2->addStretch();
        gl->addLayout(row1); gl->addLayout(row2);
        connect(d->fontFamily, &QFontComboBox::currentFontChanged, this,
                [this](const QFont& f){ emit fontFamilyChanged(f.family()); });
        connect(d->fontSize, qOverload<int>(&QSpinBox::valueChanged), this, &RibbonWidget::fontSizeChanged);
        connect(d->btnBold,      &QToolButton::toggled, this, &RibbonWidget::boldToggled);
        connect(d->btnItalic,    &QToolButton::toggled, this, &RibbonWidget::italicToggled);
        connect(d->btnUnderline, &QToolButton::toggled, this, &RibbonWidget::underlineToggled);
        connect(d->btnTextColor, &ColorButton::colorChanged, this, &RibbonWidget::textColorChanged);
        connect(d->btnFillColor, &ColorButton::colorChanged, this, &RibbonWidget::fillColorChanged);
        rowLayout->addWidget(grp); rowLayout->addWidget(vSep());
    }

    // ALIGNMENT
    {
        auto* grp = new QGroupBox("Alignment");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,2,4,14); gl->setSpacing(3);
        auto* row1 = new QHBoxLayout; row1->setSpacing(1);
        d->btnVTop = makeBtn(Icons::vTop(), "Align Top",    true);
        d->btnVMid = makeBtn(Icons::vMid(), "Align Middle", true);
        d->btnVBot = makeBtn(Icons::vBot(), "Align Bottom", true);
        d->btnWrap = makeBtn(Icons::wrap(), "Wrap Text",    true);
        d->btnMerge= makeBtn(Icons::merge(),"Merge Cells");
        row1->addWidget(d->btnVTop); row1->addWidget(d->btnVMid); row1->addWidget(d->btnVBot);
        row1->addSpacing(3); row1->addWidget(d->btnWrap); row1->addWidget(d->btnMerge);
        row1->addStretch();
        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        d->btnAlignL = makeBtn(Icons::alignL(), "Align Left",  true);
        d->btnAlignC = makeBtn(Icons::alignC(), "Center",      true);
        d->btnAlignR = makeBtn(Icons::alignR(), "Align Right", true);
        row2->addWidget(d->btnAlignL); row2->addWidget(d->btnAlignC);
        row2->addWidget(d->btnAlignR); row2->addStretch();
        gl->addLayout(row1); gl->addLayout(row2);
        connect(d->btnAlignL,&QToolButton::clicked,this,[this]{ emit hAlignChanged(0); });
        connect(d->btnAlignC,&QToolButton::clicked,this,[this]{ emit hAlignChanged(1); });
        connect(d->btnAlignR,&QToolButton::clicked,this,[this]{ emit hAlignChanged(2); });
        connect(d->btnVTop,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(0); });
        connect(d->btnVMid,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(1); });
        connect(d->btnVBot,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(2); });
        connect(d->btnWrap,  &QToolButton::toggled, this, &RibbonWidget::wrapTextToggled);
        connect(d->btnMerge, &QToolButton::clicked, this, &RibbonWidget::mergeCellsRequested);
        rowLayout->addWidget(grp); rowLayout->addWidget(vSep());
    }

    // NUMBER
    {
        auto* grp = new QGroupBox("Number");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,2,4,14); gl->setSpacing(3);
        d->numFormat = new QComboBox;
        d->numFormat->addItems({"General","Number","Currency","Percentage","Scientific"});
        d->numFormat->setFixedHeight(24);
        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        auto* btnCurr = makeBtn(Icons::currency(), "Currency Style");
        auto* btnPct  = makeBtn(Icons::percent(),  "Percent Style");
        d->btnDecInc  = makeBtn(Icons::decInc(),   "Increase Decimal");
        d->btnDecDec  = makeBtn(Icons::decDec(),   "Decrease Decimal");
        row2->addWidget(btnCurr); row2->addWidget(btnPct);
        row2->addSpacing(3); row2->addWidget(d->btnDecInc); row2->addWidget(d->btnDecDec);
        row2->addStretch();
        gl->addWidget(d->numFormat); gl->addLayout(row2);
        connect(d->numFormat,qOverload<int>(&QComboBox::currentIndexChanged),this,&RibbonWidget::numberFormatChanged);
        connect(btnCurr,&QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(2); emit numberFormatChanged(2); });
        connect(btnPct, &QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(3); emit numberFormatChanged(3); });
        connect(d->btnDecInc,&QToolButton::clicked,this,&RibbonWidget::increaseDecimalRequested);
        connect(d->btnDecDec,&QToolButton::clicked,this,&RibbonWidget::decreaseDecimalRequested);
        rowLayout->addWidget(grp); rowLayout->addWidget(vSep());
    }

    // CELLS
    {
        auto* grp = new QGroupBox("Cells");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,2,4,14); gl->setSpacing(3);
        auto* btnInsRow = makeBtn(Icons::insRow(), "Insert Row");
        auto* btnDelRow = makeBtn(Icons::delRow(), "Delete Row");
        auto* btnInsCol = makeBtn(Icons::insCol(), "Insert Column");
        auto* btnDelCol = makeBtn(Icons::delCol(), "Delete Column");
        auto* btnFmtCel = makeBtn(Icons::fmtCell(),"Format Cells");
        auto* row1 = new QHBoxLayout; row1->setSpacing(1);
        row1->addWidget(btnInsRow); row1->addWidget(btnDelRow); row1->addWidget(btnFmtCel); row1->addStretch();
        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        row2->addWidget(btnInsCol); row2->addWidget(btnDelCol); row2->addStretch();
        gl->addLayout(row1); gl->addLayout(row2);
        connect(btnInsRow,&QToolButton::clicked,this,&RibbonWidget::insertRowRequested);
        connect(btnDelRow,&QToolButton::clicked,this,&RibbonWidget::deleteRowRequested);
        connect(btnInsCol,&QToolButton::clicked,this,&RibbonWidget::insertColumnRequested);
        connect(btnDelCol,&QToolButton::clicked,this,&RibbonWidget::deleteColumnRequested);
        connect(btnFmtCel,&QToolButton::clicked,this,&RibbonWidget::formatCellsRequested);
        rowLayout->addWidget(grp); rowLayout->addWidget(vSep());
    }

    // EDITING
    {
        auto* grp = new QGroupBox("Editing");
        auto* gl  = new QVBoxLayout(grp);
        gl->setContentsMargins(4,2,4,14); gl->setSpacing(3);
        auto* btnSum    = makeBtn(Icons::autosum(), "AutoSum");
        auto* btnSortA  = makeBtn(Icons::sortAsc(), "Sort A to Z");
        auto* btnSortD  = makeBtn(Icons::sortDesc(),"Sort Z to A");
        auto* btnFilter = makeBtn(Icons::filter(),  "Filter");
        auto* btnFind   = makeBtn(Icons::find(),    "Find & Replace (Ctrl+H)");
        auto* row1 = new QHBoxLayout; row1->setSpacing(1);
        row1->addWidget(btnSum); row1->addWidget(btnSortA); row1->addWidget(btnSortD); row1->addStretch();
        auto* row2 = new QHBoxLayout; row2->setSpacing(1);
        row2->addWidget(btnFilter); row2->addWidget(btnFind); row2->addStretch();
        gl->addLayout(row1); gl->addLayout(row2);
        connect(btnSum,   &QToolButton::clicked,this,&RibbonWidget::autoSumRequested);
        connect(btnSortA, &QToolButton::clicked,this,&RibbonWidget::sortAscRequested);
        connect(btnSortD, &QToolButton::clicked,this,&RibbonWidget::sortDescRequested);
        connect(btnFilter,&QToolButton::clicked,this,&RibbonWidget::filterRequested);
        connect(btnFind,  &QToolButton::clicked,this,&RibbonWidget::findReplaceRequested);
        rowLayout->addWidget(grp);
    }

    rowLayout->addStretch();
    tabs->addTab(homeTab,     "Home");
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
#include "RibbonWidget.moc"
