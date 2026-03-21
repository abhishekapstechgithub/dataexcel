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
#include <QMenu>
#include <QWidgetAction>
#include <functional>

// ═══════════════════════════════════════════════════════════════════════════════
//  ICON FACTORY — all icons drawn with pure QPainter, zero dependencies
// ═══════════════════════════════════════════════════════════════════════════════
static QIcon makeIcon(std::function<void(QPainter&, int)> draw, int sz = 20) {
    QPixmap pm(sz, sz);
    pm.fill(Qt::transparent);
    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing);
    draw(p, sz);
    return QIcon(pm);
}

static QPen darkPen(qreal w = 1.6) {
    return QPen(QColor("#404040"), w, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
}
static QPen greenPen(qreal w = 1.8) {
    return QPen(QColor("#1e7145"), w, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
}

namespace Ic {

static QIcon paste() { return makeIcon([](QPainter& p, int){ 
    p.setPen(darkPen()); p.setBrush(QColor("#f0f0f0"));
    p.drawRoundedRect(3,5,14,13,1,1);
    p.setBrush(Qt::white); p.drawRoundedRect(6,2,8,5,1,1);
    p.setPen(darkPen(1.2)); p.drawLine(6,10,14,10); p.drawLine(6,13,14,13);
}); }

static QIcon cut() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5));
    p.drawEllipse(2,13,5,5); p.drawEllipse(13,13,5,5);
    p.drawLine(4,13,10,7); p.drawLine(16,13,10,7);
    p.drawLine(4,3,10,7);  p.drawLine(16,3,10,7);
}); }

static QIcon copy() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen()); p.setBrush(QColor("#e8e8e8"));
    p.drawRoundedRect(6,6,12,12,1,1);
    p.setBrush(Qt::white); p.drawRoundedRect(2,2,12,12,1,1);
    p.setBrush(Qt::NoBrush); p.drawRoundedRect(2,2,12,12,1,1);
}); }

static QIcon fmtpaint() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5));
    p.setBrush(QColor("#f5c518")); p.drawRoundedRect(2,2,10,8,1,1);
    p.setBrush(Qt::NoBrush);
    p.setPen(greenPen(2)); p.drawLine(12,6,18,12); p.drawLine(12,12,18,6);
}); }

static QIcon bold() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",13,QFont::Black); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"B");
}); }

static QIcon italic() { return makeIcon([](QPainter& p, int s){
    QFont f("Times New Roman",13,QFont::Normal,true); p.setFont(f);
    p.setPen(QColor("#222")); p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"I");
}); }

static QIcon underline() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",12,QFont::Bold); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,-2,s,s),Qt::AlignCenter,"U");
    p.setPen(QPen(QColor("#222"),2.5,Qt::SolidLine,Qt::RoundCap));
    p.drawLine(3,18,17,18);
}); }

static QIcon fontcolor() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",11,QFont::Bold); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,-2,s,s),Qt::AlignCenter,"A");
    p.setBrush(QColor("#e84040")); p.setPen(Qt::NoPen); p.drawRect(2,16,16,3);
}); }

static QIcon fillcolor() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5)); p.setBrush(QColor("#444"));
    QPolygon poly; poly<<QPoint(5,13)<<QPoint(10,3)<<QPoint(15,13)<<QPoint(12,13)<<QPoint(12,15)<<QPoint(8,15)<<QPoint(8,13);
    p.drawPolygon(poly);
    p.setBrush(QColor("#f5c518")); p.setPen(Qt::NoPen); p.drawRect(2,16,16,3);
}); }

static QIcon borders() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5));
    p.drawRect(3,3,14,14);
    p.setPen(darkPen(0.8));
    p.drawLine(3,10,17,10); p.drawLine(10,3,10,17);
}); }

static QIcon alignL() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8));
    p.drawLine(2,5,18,5); p.drawLine(2,9,13,9);
    p.drawLine(2,13,18,13); p.drawLine(2,17,13,17);
}); }

static QIcon alignC() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8));
    p.drawLine(2,5,18,5); p.drawLine(5,9,15,9);
    p.drawLine(2,13,18,13); p.drawLine(5,17,15,17);
}); }

static QIcon alignR() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8));
    p.drawLine(2,5,18,5); p.drawLine(7,9,18,9);
    p.drawLine(2,13,18,13); p.drawLine(7,17,18,17);
}); }

static QIcon vTop() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8)); p.drawLine(2,3,18,3);
    p.drawLine(7,5,7,17); p.drawLine(13,5,13,17); p.drawLine(7,11,13,11);
}); }

static QIcon vMid() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8)); p.drawLine(2,10,18,10);
    p.drawLine(7,3,7,17); p.drawLine(13,3,13,17);
}); }

static QIcon vBot() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8)); p.drawLine(2,17,18,17);
    p.drawLine(7,3,7,15); p.drawLine(13,3,13,15); p.drawLine(7,9,13,9);
}); }

static QIcon wrap() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5));
    p.drawLine(2,5,18,5); p.drawLine(2,10,13,10);
    p.drawArc(10,7,8,8,0,180*16);
    p.drawLine(2,15,10,15);
    QPolygon arr; arr<<QPoint(5,11)<<QPoint(2,14)<<QPoint(5,17);
    p.drawPolyline(arr);
}); }

static QIcon merge() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5));
    p.drawRect(2,2,16,16); p.drawLine(2,10,18,10); p.drawLine(10,2,10,10);
    p.setPen(greenPen(2));
    QPolygon a; a<<QPoint(4,14)<<QPoint(7,12)<<QPoint(7,16); p.drawPolygon(a);
    QPolygon b; b<<QPoint(16,14)<<QPoint(13,12)<<QPoint(13,16); p.drawPolygon(b);
    p.drawLine(5,14,15,14);
}); }

static QIcon orientation() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5));
    p.save(); p.translate(10,10); p.rotate(-45);
    p.drawLine(-7,0,7,0);
    QPolygon arr; arr<<QPoint(5,-2)<<QPoint(7,0)<<QPoint(5,2); p.drawPolygon(arr);
    p.restore();
    p.drawLine(2,16,18,16);
}); }

static QIcon currency() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",12); p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"$");
}); }

static QIcon percent() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",11); p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"%");
}); }

static QIcon thousands() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",9); p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"000");
}); }

static QIcon decInc() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",9); p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(1,14,".0"); p.setPen(greenPen(1.5));
    p.drawText(12,10,"+");
}); }

static QIcon decDec() { return makeIcon([](QPainter& p, int s){
    QFont f("Arial",9); p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(1,14,".0"); p.setPen(QPen(QColor("#c0392b"),1.5));
    p.drawText(12,10,"-");
}); }

static QIcon condFmt() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen()); p.setBrush(QColor("#fff3cd"));
    p.drawRoundedRect(2,2,16,16,2,2);
    p.setPen(QPen(QColor("#f59e0b"),2)); p.drawLine(10,5,10,12); p.drawLine(10,14,10,15);
}); }

static QIcon tableStyle() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.2));
    for(int i=0;i<3;i++) for(int j=0;j<3;j++){
        QColor c = (i==0||j==0) ? QColor("#1e7145") : (i+j)%2==0 ? QColor("#d1e8d8") : Qt::white;
        p.setBrush(c); p.setPen(Qt::NoPen);
        p.drawRect(2+i*6,2+j*6,5,5);
    }
    p.setPen(darkPen(0.8)); p.drawRect(2,2,16,16);
}); }

static QIcon rowsAndCols() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5));
    p.drawRect(2,2,16,6); p.drawRect(2,10,16,8);
    p.drawLine(10,2,10,8); p.drawLine(10,10,10,18);
}); }

static QIcon sheets() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen()); p.setBrush(Qt::white);
    p.drawRoundedRect(4,4,14,14,2,2);
    p.setBrush(QColor("#e8f4ed")); p.drawRoundedRect(2,2,14,14,2,2);
    p.setBrush(Qt::NoBrush); p.drawRoundedRect(2,2,14,14,2,2);
    p.setPen(darkPen(0.8)); p.drawLine(6,6,12,6); p.drawLine(6,9,12,9);
}); }

static QIcon insRow() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen()); p.drawRoundedRect(2,2,16,6,1,1);
    p.drawRoundedRect(2,12,16,6,1,1);
    p.setPen(greenPen(2)); p.drawLine(10,10,10,20); p.drawLine(7,15,13,15);
}); }

static QIcon delRow() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen()); p.drawRoundedRect(2,2,16,6,1,1);
    p.drawRoundedRect(2,12,16,6,1,1);
    p.setPen(QPen(QColor("#c0392b"),2,Qt::SolidLine,Qt::RoundCap)); p.drawLine(7,15,13,15);
}); }

static QIcon fmtCell() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5)); p.drawRoundedRect(2,2,16,16,2,2);
    p.drawLine(2,8,18,8); p.drawLine(8,2,8,8);
}); }

static QIcon autosum() { return makeIcon([](QPainter& p, int s){
    QFont f("Times New Roman",15); p.setFont(f); p.setPen(QColor("#1e7145"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"\u03A3");
}); }

static QIcon fill() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.5)); p.setBrush(QColor("#1e7145"));
    QPolygon poly; poly<<QPoint(4,3)<<QPoint(16,3)<<QPoint(16,15)<<QPoint(10,18)<<QPoint(4,15);
    p.drawPolygon(poly);
    p.setBrush(Qt::white); p.setPen(Qt::NoPen);
    p.drawRect(6,5,8,8);
}); }

static QIcon sortAsc() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8));
    p.drawLine(2,5,10,5); p.drawLine(2,10,14,10); p.drawLine(2,15,18,15);
    p.setPen(greenPen(2)); p.drawLine(17,3,17,11);
    QPolygon arr; arr<<QPoint(14,7)<<QPoint(17,4)<<QPoint(20,7); p.drawPolyline(arr);
}); }

static QIcon sortDesc() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8));
    p.drawLine(2,5,18,5); p.drawLine(2,10,14,10); p.drawLine(2,15,10,15);
    p.setPen(greenPen(2)); p.drawLine(17,4,17,12);
    QPolygon arr; arr<<QPoint(14,9)<<QPoint(17,12)<<QPoint(20,9); p.drawPolyline(arr);
}); }

static QIcon filter() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8));
    QPolygon poly; poly<<QPoint(2,3)<<QPoint(18,3)<<QPoint(12,10)<<QPoint(12,17)<<QPoint(8,17)<<QPoint(8,10)<<QPoint(2,3);
    p.drawPolyline(poly);
}); }

static QIcon findIcon() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen(1.8)); p.drawEllipse(2,2,12,12);
    p.setPen(QPen(QColor("#444"),2.5,Qt::SolidLine,Qt::RoundCap)); p.drawLine(12,12,18,18);
}); }

static QIcon freezePanes() { return makeIcon([](QPainter& p, int){
    p.setPen(darkPen());
    p.drawRect(2,2,16,16);
    p.setPen(QPen(QColor("#1e7145"),2,Qt::SolidLine,Qt::RoundCap));
    p.drawLine(2,8,18,8); p.drawLine(9,2,9,18);
}); }

// Dropdown arrow (small, for split buttons)
static QIcon dropArrow() { return makeIcon([](QPainter& p, int s){
    p.setPen(QPen(QColor("#555"),1.5,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin));
    p.setBrush(QColor("#555"));
    QPolygon arr; arr<<QPoint(s/2-3,s/2-1)<<QPoint(s/2+3,s/2-1)<<QPoint(s/2,s/2+2);
    p.drawPolygon(arr);
}, 10); }

} // namespace Ic

// ═══════════════════════════════════════════════════════════════════════════════
//  BUTTON FACTORIES
// ═══════════════════════════════════════════════════════════════════════════════

// Large button: icon on top, label below (like WPS Format Painter, Paste)
static QToolButton* makeLargeBtn(const QIcon& icon, const QString& label,
                                  const QString& tip, int w=54, int h=62)
{
    auto* btn = new QToolButton;
    btn->setIcon(icon);
    btn->setIconSize(QSize(28,28));
    btn->setToolTip(tip);
    btn->setText(label);
    btn->setFixedSize(w, h);
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    btn->setAutoRaise(true);
    return btn;
}

// Medium button: icon + text side by side (for Rows/Columns, Sheets etc)
static QToolButton* makeMedBtn(const QIcon& icon, const QString& label,
                                const QString& tip, int w=90, int h=30)
{
    auto* btn = new QToolButton;
    btn->setIcon(icon);
    btn->setIconSize(QSize(18,18));
    btn->setToolTip(tip);
    btn->setText("  "+label);
    btn->setFixedSize(w, h);
    btn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    btn->setAutoRaise(true);
    return btn;
}

// Medium button with dropdown arrow
static QToolButton* makeMedDropBtn(const QIcon& icon, const QString& label,
                                    const QString& tip, int w=100, int h=30)
{
    auto* btn = makeMedBtn(icon, label, tip, w, h);
    btn->setPopupMode(QToolButton::MenuButtonPopup);
    return btn;
}

// Small icon-only button
static QToolButton* makeSmallBtn(const QIcon& icon, const QString& tip,
                                  bool checkable=false)
{
    auto* btn = new QToolButton;
    btn->setIcon(icon);
    btn->setIconSize(QSize(18,18));
    btn->setToolTip(tip);
    btn->setFixedSize(28,28);
    btn->setCheckable(checkable);
    btn->setToolButtonStyle(Qt::ToolButtonIconOnly);
    btn->setAutoRaise(true);
    return btn;
}

// Vertical separator
static QFrame* vSep(int h=58) {
    auto* f = new QFrame;
    f->setFrameShape(QFrame::VLine);
    f->setFixedWidth(1); f->setFixedHeight(h);
    f->setStyleSheet("color: #d4d4d4;");
    return f;
}

// Group label at the bottom
static QLabel* groupLabel(const QString& text) {
    auto* l = new QLabel(text);
    l->setAlignment(Qt::AlignCenter);
    l->setStyleSheet("font-size: 10px; color: #888; padding: 0;");
    l->setFixedHeight(14);
    return l;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  RibbonWidget::Impl
// ═══════════════════════════════════════════════════════════════════════════════
struct RibbonWidget::Impl {
    // Font group
    QFontComboBox* fontFamily  { nullptr };
    QSpinBox*      fontSize    { nullptr };
    QToolButton*   btnBold     { nullptr };
    QToolButton*   btnItalic   { nullptr };
    QToolButton*   btnUnderline{ nullptr };
    ColorButton*   btnTextColor{ nullptr };
    ColorButton*   btnFillColor{ nullptr };
    QToolButton*   btnBorders  { nullptr };
    // Alignment group
    QToolButton*   btnAlignL   { nullptr };
    QToolButton*   btnAlignC   { nullptr };
    QToolButton*   btnAlignR   { nullptr };
    QToolButton*   btnVTop     { nullptr };
    QToolButton*   btnVMid     { nullptr };
    QToolButton*   btnVBot     { nullptr };
    QToolButton*   btnWrap     { nullptr };
    QToolButton*   btnMerge    { nullptr };
    // Number group
    QComboBox*     numFormat   { nullptr };
    QToolButton*   btnDecInc   { nullptr };
    QToolButton*   btnDecDec   { nullptr };
};

// ═══════════════════════════════════════════════════════════════════════════════
//  CONSTRUCTION
// ═══════════════════════════════════════════════════════════════════════════════
RibbonWidget::RibbonWidget(QWidget* parent)
    : QWidget(parent), d(new Impl)
{
    setFixedHeight(130);

    // ── Master stylesheet ────────────────────────────────────────────────────
    setStyleSheet(R"(
        QWidget { background: #ffffff; font-family: "Segoe UI", Arial, sans-serif; }

        /* Tab bar */
        QTabWidget::pane { border: none; background: #ffffff; margin-top: -1px; }
        QTabBar::tab {
            padding: 5px 18px 4px;
            font-size: 12px;
            font-weight: 500;
            color: #555;
            background: transparent;
            border: none;
            border-bottom: 3px solid transparent;
            min-width: 50px;
        }
        QTabBar::tab:selected { color: #1a6b35; border-bottom: 3px solid #1e7145; font-weight: 600; }
        QTabBar::tab:hover:!selected { color: #1a6b35; background: #f0f9f4; }

        /* All tool buttons */
        QToolButton {
            border: 1px solid transparent;
            border-radius: 3px;
            background: transparent;
            color: #333;
        }
        QToolButton:hover  { background: #e8f5ee; border-color: #b8d9c4; }
        QToolButton:pressed, QToolButton:checked { background: #c8e6d4; border-color: #1e7145; }

        /* Large buttons (Paste, Format Painter) */
        QToolButton[large="true"] {
            font-size: 11px; padding: 2px;
        }
        QToolButton[large="true"]:hover { background: #e8f5ee; }

        /* Medium text buttons */
        QToolButton[medium="true"] {
            font-size: 11px; text-align: left; padding: 2px 4px;
        }

        /* Group box labels at bottom */
        QGroupBox {
            border: none;
            font-size: 10px;
            color: #888;
            margin-top: 0px;
            padding: 0px;
        }
        QGroupBox::title {
            subcontrol-origin: padding;
            subcontrol-position: bottom center;
            padding: 0 4px;
        }

        /* Dropdowns */
        QComboBox {
            border: 1px solid #d0d0d0;
            border-radius: 3px;
            background: white;
            padding: 2px 6px;
            font-size: 12px;
            min-width: 40px;
            selection-background-color: #c8e6d4;
        }
        QComboBox:hover { border-color: #1e7145; }
        QComboBox::drop-down { border: none; width: 18px; }
        QFontComboBox {
            border: 1px solid #d0d0d0;
            border-radius: 3px;
            background: white;
            font-size: 12px;
            min-width: 130px;
        }
        QFontComboBox:hover { border-color: #1e7145; }
        QSpinBox {
            border: 1px solid #d0d0d0;
            border-radius: 3px;
            background: white;
            font-size: 12px;
            min-width: 48px;
        }
        QSpinBox:hover { border-color: #1e7145; }

        /* Thin rule under ribbon */
        QFrame[ribbon_sep="true"] { background: #e0e0e0; }
    )");

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(0,0,0,0);
    mainLayout->setSpacing(0);

    // ── Tab widget ────────────────────────────────────────────────────────────
    auto* tabs = new QTabWidget;
    tabs->setTabPosition(QTabWidget::North);
    tabs->setDocumentMode(true);
    tabs->setElideMode(Qt::ElideNone);

    // ════════════════════════════════════════════════════════════════════════
    //  HOME TAB
    // ════════════════════════════════════════════════════════════════════════
    auto* homeTab = new QWidget;
    homeTab->setStyleSheet("background:#ffffff;");
    auto* homeRow = new QHBoxLayout(homeTab);
    homeRow->setContentsMargins(6,2,6,0);
    homeRow->setSpacing(0);

    auto addSep = [&]{ homeRow->addWidget(vSep()); };
    auto addSpace = [&](int px=4){ homeRow->addSpacing(px); };

    // ── CLIPBOARD ─────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(0);

        auto* row = new QHBoxLayout;
        row->setSpacing(2);

        auto* btnPaste = makeLargeBtn(Ic::paste(), "Paste", "Paste (Ctrl+V)", 56, 62);
        btnPaste->setProperty("large", true);

        auto* vstack = new QVBoxLayout; vstack->setSpacing(1);
        auto* btnCut  = makeSmallBtn(Ic::cut(),      "Cut (Ctrl+X)");
        auto* btnCopy = makeSmallBtn(Ic::copy(),     "Copy (Ctrl+C)");
        auto* btnFmt  = makeSmallBtn(Ic::fmtpaint(), "Format Painter");
        vstack->addWidget(btnCut); vstack->addWidget(btnCopy); vstack->addWidget(btnFmt);

        row->addWidget(btnPaste);
        row->addSpacing(2);
        row->addLayout(vstack);

        col->addLayout(row);
        col->addWidget(groupLabel("Clipboard"));
        homeRow->addWidget(grp);
        addSep(); addSpace();

        connect(btnPaste,&QToolButton::clicked,this,&RibbonWidget::pasteRequested);
        connect(btnCut,  &QToolButton::clicked,this,&RibbonWidget::cutRequested);
        connect(btnCopy, &QToolButton::clicked,this,&RibbonWidget::copyRequested);
        connect(btnFmt,  &QToolButton::clicked,this,&RibbonWidget::formatPainterRequested);
    }

    // ── FONT ──────────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        // Row 1: font family + size + grow/shrink
        auto* r1 = new QHBoxLayout; r1->setSpacing(3);
        d->fontFamily = new QFontComboBox;
        d->fontFamily->setCurrentFont(QFont("Calibri"));
        d->fontFamily->setFixedHeight(24); d->fontFamily->setMaximumWidth(150);
        d->fontSize = new QSpinBox;
        d->fontSize->setRange(6,96); d->fontSize->setValue(11); d->fontSize->setFixedHeight(24);
        d->fontSize->setFixedWidth(52);

        auto* btnGrow   = makeSmallBtn(makeIcon([](QPainter& p,int s){
            QFont f("Arial",11,QFont::Bold); p.setFont(f); p.setPen(QColor("#444"));
            p.drawText(QRect(0,-1,s,s),Qt::AlignCenter,"A");
            p.setFont(QFont("Arial",7)); p.setPen(greenPen(1.2));
            p.drawText(QRect(s/2,s/2-2,s/2+2,s/2+2),Qt::AlignCenter,"+");
        }), "Increase Font Size");
        auto* btnShrink = makeSmallBtn(makeIcon([](QPainter& p,int s){
            QFont f("Arial",9); p.setFont(f); p.setPen(QColor("#444"));
            p.drawText(QRect(0,-1,s,s),Qt::AlignCenter,"A");
            p.setFont(QFont("Arial",7)); p.setPen(QPen(QColor("#c0392b"),1.2));
            p.drawText(QRect(s/2,s/2-2,s/2+2,s/2+2),Qt::AlignCenter,"-");
        }), "Decrease Font Size");

        r1->addWidget(d->fontFamily,1);
        r1->addWidget(d->fontSize);
        r1->addWidget(btnGrow);
        r1->addWidget(btnShrink);

        // Row 2: B I U | borders | A fill | color strip buttons
        auto* r2 = new QHBoxLayout; r2->setSpacing(1);
        d->btnBold      = makeSmallBtn(Ic::bold(),      "Bold (Ctrl+B)",      true);
        d->btnItalic    = makeSmallBtn(Ic::italic(),    "Italic (Ctrl+I)",    true);
        d->btnUnderline = makeSmallBtn(Ic::underline(), "Underline (Ctrl+U)", true);
        d->btnBorders   = makeSmallBtn(Ic::borders(),   "Borders");

        // Font color button with color strip
        d->btnTextColor = new ColorButton("", grp);
        d->btnTextColor->setToolTip("Font Color");
        d->btnTextColor->setFixedSize(28,28);

        // Fill color button
        d->btnFillColor = new ColorButton("", grp);
        d->btnFillColor->setColor(Qt::yellow);
        d->btnFillColor->setToolTip("Highlight/Fill Color");
        d->btnFillColor->setFixedSize(28,28);

        r2->addWidget(d->btnBold); r2->addWidget(d->btnItalic);
        r2->addWidget(d->btnUnderline); r2->addSpacing(3);
        r2->addWidget(d->btnBorders); r2->addSpacing(3);
        r2->addWidget(d->btnTextColor); r2->addWidget(d->btnFillColor);
        r2->addStretch();

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Font"));
        homeRow->addWidget(grp);
        addSep(); addSpace();

        connect(d->fontFamily,&QFontComboBox::currentFontChanged,this,
                [this](const QFont& f){ emit fontFamilyChanged(f.family()); });
        connect(d->fontSize,qOverload<int>(&QSpinBox::valueChanged),this,&RibbonWidget::fontSizeChanged);
        connect(d->btnBold,      &QToolButton::toggled,this,&RibbonWidget::boldToggled);
        connect(d->btnItalic,    &QToolButton::toggled,this,&RibbonWidget::italicToggled);
        connect(d->btnUnderline, &QToolButton::toggled,this,&RibbonWidget::underlineToggled);
        connect(d->btnTextColor, &ColorButton::colorChanged,this,&RibbonWidget::textColorChanged);
        connect(d->btnFillColor, &ColorButton::colorChanged,this,&RibbonWidget::fillColorChanged);
    }

    // ── ALIGNMENT ─────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        // Row 1: vertical align | orientation | wrap | merge
        auto* r1 = new QHBoxLayout; r1->setSpacing(1);
        d->btnVTop  = makeSmallBtn(Ic::vTop(),       "Align Top",    true);
        d->btnVMid  = makeSmallBtn(Ic::vMid(),       "Align Middle", true);
        d->btnVBot  = makeSmallBtn(Ic::vBot(),       "Align Bottom", true);
        auto* btnOri = makeSmallBtn(Ic::orientation(),"Orientation");
        d->btnWrap  = makeSmallBtn(Ic::wrap(),       "Wrap Text",    true);
        d->btnMerge = makeSmallBtn(Ic::merge(),      "Merge and Center");
        r1->addWidget(d->btnVTop); r1->addWidget(d->btnVMid); r1->addWidget(d->btnVBot);
        r1->addSpacing(3);
        r1->addWidget(btnOri); r1->addWidget(d->btnWrap); r1->addWidget(d->btnMerge);
        r1->addStretch();

        // Row 2: horizontal align | indent
        auto* r2 = new QHBoxLayout; r2->setSpacing(1);
        d->btnAlignL = makeSmallBtn(Ic::alignL(), "Align Left",   true);
        d->btnAlignC = makeSmallBtn(Ic::alignC(), "Center",       true);
        d->btnAlignR = makeSmallBtn(Ic::alignR(), "Align Right",  true);
        auto* btnIndL = makeSmallBtn(makeIcon([](QPainter& p,int){
            p.setPen(darkPen(1.5));
            p.drawLine(2,5,18,5); p.drawLine(6,9,18,9); p.drawLine(2,13,18,13); p.drawLine(6,17,18,17);
            p.setPen(greenPen(2)); QPolygon a; a<<QPoint(2,9)<<QPoint(2,17)<<QPoint(5,13); p.drawPolygon(a);
        }), "Increase Indent");
        auto* btnIndR = makeSmallBtn(makeIcon([](QPainter& p,int){
            p.setPen(darkPen(1.5));
            p.drawLine(2,5,18,5); p.drawLine(2,9,14,9); p.drawLine(2,13,18,13); p.drawLine(2,17,14,17);
            p.setPen(greenPen(2)); QPolygon a; a<<QPoint(18,9)<<QPoint(18,17)<<QPoint(15,13); p.drawPolygon(a);
        }), "Decrease Indent");
        r2->addWidget(d->btnAlignL); r2->addWidget(d->btnAlignC); r2->addWidget(d->btnAlignR);
        r2->addSpacing(3);
        r2->addWidget(btnIndL); r2->addWidget(btnIndR);
        r2->addStretch();

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Alignment"));
        homeRow->addWidget(grp);
        addSep(); addSpace();

        connect(d->btnAlignL,&QToolButton::clicked,this,[this]{ emit hAlignChanged(0); });
        connect(d->btnAlignC,&QToolButton::clicked,this,[this]{ emit hAlignChanged(1); });
        connect(d->btnAlignR,&QToolButton::clicked,this,[this]{ emit hAlignChanged(2); });
        connect(d->btnVTop,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(0); });
        connect(d->btnVMid,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(1); });
        connect(d->btnVBot,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(2); });
        connect(d->btnWrap,  &QToolButton::toggled,this,&RibbonWidget::wrapTextToggled);
        connect(d->btnMerge, &QToolButton::clicked,this,&RibbonWidget::mergeCellsRequested);
    }

    // ── NUMBER ────────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        // Row 1: format dropdown
        d->numFormat = new QComboBox;
        d->numFormat->addItems({"General","Number","Currency","Accounting",
                                "Short Date","Long Date","Time","Percentage",
                                "Fraction","Scientific","Text"});
        d->numFormat->setFixedHeight(24); d->numFormat->setFixedWidth(130);

        // Row 2: $ % 000 | .0+ .0-
        auto* r2 = new QHBoxLayout; r2->setSpacing(1);
        auto* btnCurr = makeSmallBtn(Ic::currency(),  "Currency ($)");
        auto* btnPct  = makeSmallBtn(Ic::percent(),   "Percent Style (%)");
        auto* btnThou = makeSmallBtn(Ic::thousands(),  "Thousands Separator");
        d->btnDecInc  = makeSmallBtn(Ic::decInc(),    "Increase Decimal");
        d->btnDecDec  = makeSmallBtn(Ic::decDec(),    "Decrease Decimal");
        r2->addWidget(btnCurr); r2->addWidget(btnPct); r2->addWidget(btnThou);
        r2->addSpacing(4);
        r2->addWidget(d->btnDecInc); r2->addWidget(d->btnDecDec);
        r2->addStretch();

        col->addWidget(d->numFormat);
        col->addLayout(r2);
        col->addWidget(groupLabel("Number"));
        homeRow->addWidget(grp);
        addSep(); addSpace();

        connect(d->numFormat,qOverload<int>(&QComboBox::currentIndexChanged),this,&RibbonWidget::numberFormatChanged);
        connect(btnCurr,&QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(2); emit numberFormatChanged(2); });
        connect(btnPct, &QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(7); emit numberFormatChanged(7); });
        connect(d->btnDecInc,&QToolButton::clicked,this,&RibbonWidget::increaseDecimalRequested);
        connect(d->btnDecDec,&QToolButton::clicked,this,&RibbonWidget::decreaseDecimalRequested);
    }

    // ── STYLES ────────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        auto* r1 = new QHBoxLayout; r1->setSpacing(2);
        auto* btnCond  = makeMedDropBtn(Ic::condFmt(),   "Conditional\nFormatting","Conditional Formatting",110,30);
        auto* btnTable = makeMedDropBtn(Ic::tableStyle(), "Table\nStyles",          "Table Styles",90,30);
        r1->addWidget(btnCond); r1->addWidget(btnTable);

        auto* r2 = new QHBoxLayout; r2->setSpacing(2);
        auto* btnRowCol = makeMedDropBtn(Ic::rowsAndCols(), "Rows and\nColumns", "Rows and Columns",110,30);
        auto* btnSheets = makeMedDropBtn(Ic::sheets(),       "Sheets",            "Sheet Operations",90,30);
        r2->addWidget(btnRowCol); r2->addWidget(btnSheets);

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Styles"));
        homeRow->addWidget(grp);
        addSep(); addSpace();
    }

    // ── CELLS ─────────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        auto* r1 = new QHBoxLayout; r1->setSpacing(2);
        auto* btnInsRow = makeSmallBtn(Ic::insRow(), "Insert Row");
        auto* btnDelRow = makeSmallBtn(Ic::delRow(), "Delete Row");
        auto* btnFmtCel = makeSmallBtn(Ic::fmtCell(),"Format Cells");
        r1->addWidget(btnInsRow); r1->addWidget(btnDelRow); r1->addWidget(btnFmtCel);

        auto* r2 = new QHBoxLayout; r2->setSpacing(2);
        auto* btnInsCol = makeSmallBtn(makeIcon([](QPainter& p,int){
            p.setPen(darkPen()); p.drawRoundedRect(2,2,6,16,1,1); p.drawRoundedRect(12,2,6,16,1,1);
            p.setPen(greenPen(2)); p.drawLine(8,5,8,15); p.drawLine(5,10,11,10);
        }), "Insert Column");
        auto* btnDelCol = makeSmallBtn(makeIcon([](QPainter& p,int){
            p.setPen(darkPen()); p.drawRoundedRect(2,2,6,16,1,1); p.drawRoundedRect(12,2,6,16,1,1);
            p.setPen(QPen(QColor("#c0392b"),2,Qt::SolidLine,Qt::RoundCap)); p.drawLine(5,10,11,10);
        }), "Delete Column");
        r2->addWidget(btnInsCol); r2->addWidget(btnDelCol);
        r2->addStretch();

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Cells"));
        homeRow->addWidget(grp);
        addSep(); addSpace();

        connect(btnInsRow,&QToolButton::clicked,this,&RibbonWidget::insertRowRequested);
        connect(btnDelRow,&QToolButton::clicked,this,&RibbonWidget::deleteRowRequested);
        connect(btnInsCol,&QToolButton::clicked,this,&RibbonWidget::insertColumnRequested);
        connect(btnDelCol,&QToolButton::clicked,this,&RibbonWidget::deleteColumnRequested);
        connect(btnFmtCel,&QToolButton::clicked,this,&RibbonWidget::formatCellsRequested);
    }

    // ── EDITING ───────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        auto* r1 = new QHBoxLayout; r1->setSpacing(2);
        auto* btnSum    = makeLargeBtn(Ic::autosum(),  "AutoSum",  "AutoSum",  48, 62);
        auto* btnFill   = makeLargeBtn(Ic::fill(),     "Fill",     "Fill Down/Right", 44, 62);
        btnSum->setProperty("large",true); btnFill->setProperty("large",true);

        auto* vstack = new QVBoxLayout; vstack->setSpacing(1);
        auto* btnSortA  = makeSmallBtn(Ic::sortAsc(),  "Sort A to Z");
        auto* btnSortD  = makeSmallBtn(Ic::sortDesc(), "Sort Z to A");
        auto* btnFilter = makeSmallBtn(Ic::filter(),   "Filter");
        auto* btnFind   = makeSmallBtn(Ic::findIcon(), "Find & Replace (Ctrl+H)");
        auto* btnFreeze = makeSmallBtn(Ic::freezePanes(),"Freeze Panes");

        auto* vr1 = new QHBoxLayout; vr1->setSpacing(1);
        vr1->addWidget(btnSortA); vr1->addWidget(btnSortD);
        auto* vr2 = new QHBoxLayout; vr2->setSpacing(1);
        vr2->addWidget(btnFilter); vr2->addWidget(btnFind); vr2->addWidget(btnFreeze);

        vstack->addLayout(vr1); vstack->addLayout(vr2);

        r1->addWidget(btnSum);
        r1->addWidget(btnFill);
        r1->addSpacing(3);
        r1->addLayout(vstack);

        col->addLayout(r1);
        col->addWidget(groupLabel("Editing"));
        homeRow->addWidget(grp);

        connect(btnSum,   &QToolButton::clicked,this,&RibbonWidget::autoSumRequested);
        connect(btnSortA, &QToolButton::clicked,this,&RibbonWidget::sortAscRequested);
        connect(btnSortD, &QToolButton::clicked,this,&RibbonWidget::sortDescRequested);
        connect(btnFilter,&QToolButton::clicked,this,&RibbonWidget::filterRequested);
        connect(btnFind,  &QToolButton::clicked,this,&RibbonWidget::findReplaceRequested);
    }

    homeRow->addStretch();

    // ── Tab setup ──────────────────────────────────────────────────────────
    tabs->addTab(homeTab,     "Home");
    tabs->addTab(new QWidget, "Insert");
    tabs->addTab(new QWidget, "Page Layout");
    tabs->addTab(new QWidget, "Formulas");
    tabs->addTab(new QWidget, "Data");
    tabs->addTab(new QWidget, "Review");
    tabs->addTab(new QWidget, "View");

    mainLayout->addWidget(tabs);

    // Bottom border line
    auto* bottomLine = new QFrame;
    bottomLine->setFixedHeight(1);
    bottomLine->setProperty("ribbon_sep", true);
    bottomLine->setStyleSheet("background: #e0e0e0;");
    mainLayout->addWidget(bottomLine);
}

RibbonWidget::~RibbonWidget() { delete d; }

// ═══════════════════════════════════════════════════════════════════════════════
void RibbonWidget::setFormatState(const RibbonFormatState& s) {
    QSignalBlocker b1(d->fontFamily), b2(d->fontSize),
                   b3(d->btnBold),    b4(d->btnItalic),
                   b5(d->btnUnderline), b6(d->numFormat);

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
