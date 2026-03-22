// ═══════════════════════════════════════════════════════════════════════════════
//  RibbonWidget.cpp — Full WPS-style ribbon with all Home tab groups
//  Groups: Clipboard | Font | Alignment | Number Format | Styles | Cells | Editing
//  + Format Conversion | Conditional Formatting | Cell Styles | Table Styles
//  + Sheets | Fill | AutoSum | Sort | Filter | Freeze Panes | Find
// ═══════════════════════════════════════════════════════════════════════════════
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
//  ICON FACTORY
// ═══════════════════════════════════════════════════════════════════════════════
static QIcon makeIcon(std::function<void(QPainter&, int)> draw, int sz = 20) {
    QPixmap pm(sz, sz); pm.fill(Qt::transparent);
    QPainter p(&pm); p.setRenderHint(QPainter::Antialiasing); draw(p, sz);
    return QIcon(pm);
}
static QPen darkPen(qreal w=1.6){ return QPen(QColor("#404040"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }
static QPen greenPen(qreal w=1.8){ return QPen(QColor("#1e7145"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }
static QPen redPen(qreal w=1.6){ return QPen(QColor("#c0392b"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }

namespace Ic {
static QIcon paste(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen()); p.setBrush(QColor("#e8e8e8"));
    p.drawRoundedRect(2,5,14,13,1,1);
    p.setBrush(Qt::white); p.setPen(greenPen(1.5));
    p.drawRoundedRect(5,2,10,6,1,1);
    p.setPen(darkPen(1.1)); p.drawLine(5,10,13,10); p.drawLine(5,13,13,13);
});}
static QIcon cut(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.6));
    p.drawEllipse(2,12,5,5); p.drawEllipse(13,12,5,5);
    p.drawLine(4,12,10,6); p.drawLine(16,12,10,6);
    p.drawLine(4,2,10,6); p.drawLine(16,2,10,6);
});}
static QIcon copy(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen()); p.setBrush(QColor("#e8e8e8"));
    p.drawRoundedRect(6,6,12,12,1,1);
    p.setBrush(Qt::white); p.drawRoundedRect(2,2,12,12,1,1);
});}
static QIcon fmtpaint(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(2)); p.setBrush(QColor("#1e7145"));
    p.drawRoundedRect(1,1,10,12,1,1);
    p.setPen(darkPen(1.5)); p.setBrush(Qt::NoBrush);
    p.drawLine(11,12,17,18); p.drawLine(14,9,17,6);
});}
static QIcon bold(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.65,QFont::Bold); p.setFont(f); p.setPen(QColor("#333"));
    p.drawText(QRect(1,0,s,s),Qt::AlignCenter,"B");
});}
static QIcon italic(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.65,QFont::Normal,true); p.setFont(f); p.setPen(QColor("#333"));
    p.drawText(QRect(1,0,s,s),Qt::AlignCenter,"I");
});}
static QIcon underline(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.55,QFont::Bold); p.setFont(f); p.setPen(QColor("#333"));
    p.drawText(QRect(1,-2,s,s),Qt::AlignCenter,"U");
    p.setPen(QPen(QColor("#1e7145"),2)); p.drawLine(2,s-2,s-2,s-2);
});}
static QIcon borders(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.2));
    p.drawRect(2,2,16,16);
    p.drawLine(10,2,10,18); p.drawLine(2,10,18,10);
});}
static QIcon alignL(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.5));
    p.drawLine(2,5,18,5); p.drawLine(2,9,14,9); p.drawLine(2,13,18,13); p.drawLine(2,17,14,17);
});}
static QIcon alignC(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.5));
    p.drawLine(2,5,18,5); p.drawLine(5,9,15,9); p.drawLine(2,13,18,13); p.drawLine(5,17,15,17);
});}
static QIcon alignR(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.5));
    p.drawLine(2,5,18,5); p.drawLine(6,9,18,9); p.drawLine(2,13,18,13); p.drawLine(6,17,18,17);
});}
static QIcon vTop(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(2)); p.drawLine(2,2,18,2);
    p.setPen(darkPen(1.5)); p.drawLine(6,4,10,14); p.drawLine(14,4,10,14);
    p.drawLine(6,4,14,4);
});}
static QIcon vMid(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(2)); p.drawLine(2,10,18,10);
    p.setPen(darkPen(1.5)); p.drawLine(5,3,5,17); p.drawLine(15,3,15,17);
    p.drawLine(5,3,15,3); p.drawLine(5,17,15,17);
});}
static QIcon vBot(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(2)); p.drawLine(2,18,18,18);
    p.setPen(darkPen(1.5)); p.drawLine(6,16,10,6); p.drawLine(14,16,10,6);
    p.drawLine(6,16,14,16);
});}
static QIcon wrap(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.4));
    p.drawLine(2,5,16,5); p.drawLine(2,9,18,9);
    QPainterPath path; path.moveTo(18,9); path.cubicTo(20,9,20,14,18,14);
    path.lineTo(14,14); p.drawPath(path);
    p.setBrush(QColor("#404040")); p.setPen(Qt::NoPen);
    QPolygon arr; arr<<QPoint(14,11)<<QPoint(14,17)<<QPoint(10,14); p.drawPolygon(arr);
    p.setPen(darkPen(1.4)); p.drawLine(2,13,10,13);
});}
static QIcon merge(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen()); p.setBrush(Qt::NoBrush);
    p.drawRect(2,2,7,7); p.drawRect(11,2,7,7); p.drawRect(2,11,7,7); p.drawRect(11,11,7,7);
    p.setPen(greenPen(2));
    p.drawLine(10,2,10,18); p.drawLine(2,10,18,10);
});}
static QIcon orientation(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(1.8));
    p.save(); p.translate(10,10); p.rotate(-45); p.drawLine(-7,0,7,0);
    QPolygon a; a<<QPoint(4,-3)<<QPoint(7,0)<<QPoint(4,3); p.setPen(Qt::NoPen);
    p.setBrush(QColor("#1e7145")); p.drawPolygon(a); p.restore();
    p.setPen(darkPen(1.3)); p.drawLine(2,17,18,17);
});}
static QIcon currency(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.7,QFont::Bold); p.setFont(f); p.setPen(QColor("#1e7145"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"$");
});}
static QIcon percent(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.65,QFont::Bold); p.setFont(f); p.setPen(QColor("#333"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"%");
});}
static QIcon thousands(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.5); p.setFont(f); p.setPen(QColor("#333"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"000");
});}
static QIcon decInc(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.45); p.setFont(f); p.setPen(QColor("#1e7145"));
    p.drawText(QRect(0,0,s/2+2,s),Qt::AlignVCenter|Qt::AlignRight,".0");
    p.setPen(QColor("#333")); p.drawText(QRect(s/2,0,s/2,s),Qt::AlignVCenter|Qt::AlignLeft,"0>");
});}
static QIcon decDec(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.45); p.setFont(f); p.setPen(QColor("#c0392b"));
    p.drawText(QRect(0,0,s/2+4,s),Qt::AlignVCenter|Qt::AlignRight,"<0");
    p.setPen(QColor("#333")); p.drawText(QRect(s/2+2,0,s/2,s),Qt::AlignVCenter|Qt::AlignLeft,".0");
});}
static QIcon condFmt(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.2));
    p.setBrush(QColor("#ffe8e8")); p.drawRoundedRect(2,2,16,7,1,1);
    p.setBrush(QColor("#e8ffe8")); p.drawRoundedRect(2,11,16,7,1,1);
    p.setPen(redPen(1.5)); p.drawLine(5,5,8,5); p.drawLine(10,5,15,5);
    p.setPen(greenPen(1.5)); p.drawLine(5,14,15,14);
});}
static QIcon fmtConvert(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen()); p.setBrush(QColor("#e8f5ee"));
    p.drawRoundedRect(1,1,12,16,2,2);
    p.setPen(greenPen(2)); p.drawLine(13,9,19,9);
    p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(16,6)<<QPoint(19,9)<<QPoint(16,12); p.drawPolygon(a);
});}
static QIcon tableStyle(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.1));
    p.setBrush(QColor("#1e7145")); p.drawRect(2,2,16,4);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,6,8,4);
    p.setBrush(Qt::white);        p.drawRect(10,6,8,4);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,10,8,4);
    p.setBrush(Qt::white);        p.drawRect(10,10,8,4);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,14,8,4);
    p.setBrush(Qt::white);        p.drawRect(10,14,8,4);
});}
static QIcon cellStyle(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(2)); p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(2,2,16,16,2,2);
    p.setPen(darkPen(1.2));
    p.drawLine(2,8,18,8); p.drawLine(10,2,10,18);
    p.setBrush(QColor("#c8e6d4")); p.setPen(Qt::NoPen);
    p.drawRoundedRect(3,3,6,4,1,1);
});}
static QIcon rowsAndCols(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.3));
    p.drawLine(2,6,18,6); p.drawLine(2,10,18,10); p.drawLine(2,14,18,14);
    p.setPen(greenPen(1.8));
    p.drawLine(10,2,10,18);
    p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(7,3)<<QPoint(10,0)<<QPoint(13,3); p.drawPolygon(a);
    QPolygon b; b<<QPoint(7,17)<<QPoint(10,20)<<QPoint(13,17); p.drawPolygon(b);
});}
static QIcon sheets(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen()); p.setBrush(QColor("#e8e8e8"));
    p.drawRoundedRect(2,6,18,13,1,1);
    p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(4,4,8,5,1,1);
    p.setPen(greenPen(1.5)); p.drawLine(14,3,14,8);
    p.drawLine(11,6,17,6);
});}
static QIcon autosum(){return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",s*0.65,QFont::Bold); p.setFont(f); p.setPen(QColor("#1e7145"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"Σ");
});}
static QIcon fill(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen()); p.setBrush(QColor("#e8f5ee"));
    p.drawRoundedRect(2,2,16,16,1,1);
    p.setPen(greenPen(2));
    p.drawLine(10,5,10,15); p.drawLine(5,10,15,10);
});}
static QIcon sortAsc(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.5));
    p.drawLine(2,16,8,4); p.drawLine(8,4,14,16);
    p.drawLine(4,12,12,12);
    p.setPen(darkPen(1.2));
    p.drawLine(16,5,18,5); p.drawLine(16,9,18,9); p.drawLine(16,13,18,13);
});}
static QIcon sortDesc(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.5));
    p.drawLine(2,4,8,16); p.drawLine(8,16,14,4);
    p.drawLine(4,8,12,8);
    p.setPen(darkPen(1.2));
    p.drawLine(16,5,18,5); p.drawLine(16,9,18,9); p.drawLine(16,13,18,13);
});}
static QIcon filter(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(2)); p.setBrush(Qt::NoBrush);
    p.drawLine(2,4,18,4); p.drawLine(5,9,15,9); p.drawLine(8,14,12,14);
    p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(5,18)<<QPoint(8,15)<<QPoint(8,9)<<QPoint(5,9); p.drawPolygon(a);
});}
static QIcon findIcon(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.8)); p.setBrush(Qt::NoBrush);
    p.drawEllipse(2,2,12,12); p.drawLine(11,11,18,18);
});}
static QIcon freezePanes(){return makeIcon([](QPainter&p,int){
    p.setPen(greenPen(2));
    p.drawLine(2,8,18,8); p.drawLine(10,2,10,18);
    p.setPen(darkPen(1.1));
    p.drawRect(2,2,8,6); p.drawRect(2,8,8,10); p.drawRect(10,2,8,6); p.drawRect(10,8,8,10);
});}
static QIcon insRow(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.2));
    p.drawLine(2,6,18,6); p.drawLine(2,14,18,14);
    p.setPen(greenPen(2));
    p.drawLine(10,8,10,12); p.drawLine(7,10,13,10);
});}
static QIcon delRow(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen(1.2));
    p.drawLine(2,6,18,6); p.drawLine(2,14,18,14);
    p.setPen(redPen(2));
    p.drawLine(7,10,13,10);
});}
static QIcon fmtCell(){return makeIcon([](QPainter&p,int){
    p.setPen(darkPen()); p.setBrush(Qt::white);
    p.drawRoundedRect(2,2,16,16,2,2);
    p.setPen(greenPen(1.5));
    p.drawLine(5,6,15,6); p.drawLine(5,10,15,10); p.drawLine(5,14,11,14);
});}
} // namespace Ic

// ═══════════════════════════════════════════════════════════════════════════════
//  HELPER WIDGETS
// ═══════════════════════════════════════════════════════════════════════════════
static QFrame* vSep() {
    auto* f = new QFrame; f->setFrameShape(QFrame::VLine);
    f->setFixedWidth(1); f->setFixedHeight(80);
    f->setStyleSheet("color: #e0e0e0;"); return f;
}

// Group label at the bottom of each section
static QLabel* groupLabel(const QString& text) {
    auto* lbl = new QLabel(text);
    lbl->setAlignment(Qt::AlignCenter);
    lbl->setStyleSheet("color:#888; font-size:10px; font-family:'Segoe UI',Arial;");
    lbl->setFixedHeight(14);
    return lbl;
}

// Large button (Paste / AutoSum / Fill)
static QToolButton* makeLargeBtn(const QIcon& icon, const QString& label,
                                  const QString& tip, int w=52, int h=60) {
    auto* btn = new QToolButton;
    btn->setIcon(icon); btn->setIconSize(QSize(28,28));
    btn->setText(label); btn->setToolTip(tip);
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    btn->setAutoRaise(true);
    btn->setFixedSize(w, h);
    btn->setStyleSheet(
        "QToolButton { border:1px solid transparent; border-radius:3px; "
        "  background:transparent; font-size:10px; color:#333; padding:2px; }"
        "QToolButton:hover { background:#e8f5ee; border-color:#b8d9c4; }"
        "QToolButton:pressed { background:#c8e6d4; border-color:#1e7145; }"
    );
    return btn;
}

// Small square button
static QToolButton* makeSmallBtn(const QIcon& icon, const QString& tip,
                                  bool checkable=false, int sz=28) {
    auto* btn = new QToolButton;
    btn->setIcon(icon); btn->setIconSize(QSize(sz-8, sz-8));
    btn->setToolTip(tip); btn->setAutoRaise(true);
    btn->setCheckable(checkable);
    btn->setFixedSize(sz, sz);
    btn->setStyleSheet(
        "QToolButton { border:1px solid transparent; border-radius:3px; background:transparent; }"
        "QToolButton:hover { background:#e8f5ee; border-color:#b8d9c4; }"
        "QToolButton:pressed, QToolButton:checked { background:#c8e6d4; border-color:#1e7145; }"
    );
    return btn;
}

// Medium two-line drop button (like "Conditional\nFormatting")
static QToolButton* makeMedDropBtn(const QIcon& icon, const QString& label,
                                    const QString& tip, int w=100, int h=56) {
    auto* btn = new QToolButton;
    btn->setIcon(icon); btn->setIconSize(QSize(22,22));
    btn->setText(label); btn->setToolTip(tip);
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    btn->setPopupMode(QToolButton::MenuButtonPopup);
    btn->setAutoRaise(true);
    btn->setFixedSize(w, h);
    btn->setStyleSheet(
        "QToolButton { border:1px solid transparent; border-radius:3px; "
        "  background:transparent; font-size:10px; color:#333; padding:1px; }"
        "QToolButton:hover { background:#e8f5ee; border-color:#b8d9c4; }"
        "QToolButton:pressed { background:#c8e6d4; border-color:#1e7145; }"
        "QToolButton::menu-button { width:14px; border:none; }"
    );
    return btn;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  Impl struct
// ═══════════════════════════════════════════════════════════════════════════════
struct RibbonWidget::Impl {
    QFontComboBox* fontFamily  {nullptr};
    QSpinBox*      fontSize    {nullptr};
    QToolButton*   btnBold     {nullptr};
    QToolButton*   btnItalic   {nullptr};
    QToolButton*   btnUnderline{nullptr};
    QToolButton*   btnBorders  {nullptr};
    ColorButton*   btnTextColor{nullptr};
    ColorButton*   btnFillColor{nullptr};
    QToolButton*   btnAlignL   {nullptr};
    QToolButton*   btnAlignC   {nullptr};
    QToolButton*   btnAlignR   {nullptr};
    QToolButton*   btnVTop     {nullptr};
    QToolButton*   btnVMid     {nullptr};
    QToolButton*   btnVBot     {nullptr};
    QToolButton*   btnWrap     {nullptr};
    QToolButton*   btnMerge    {nullptr};
    QComboBox*     numFormat   {nullptr};
    QToolButton*   btnDecInc   {nullptr};
    QToolButton*   btnDecDec   {nullptr};
};

// ═══════════════════════════════════════════════════════════════════════════════
//  CONSTRUCTION
// ═══════════════════════════════════════════════════════════════════════════════
RibbonWidget::RibbonWidget(QWidget* parent)
    : QWidget(parent), d(new Impl)
{
    setFixedHeight(132);
    setStyleSheet(R"(
        QWidget { background: #ffffff; font-family: "Segoe UI", Arial, sans-serif; }

        QTabWidget::pane { border: none; background: #ffffff; margin-top: -1px; }
        QTabBar::tab {
            padding: 5px 16px 4px;
            font-size: 12px;
            font-weight: 500;
            color: #555;
            background: transparent;
            border: none;
            border-bottom: 3px solid transparent;
            min-width: 46px;
        }
        QTabBar::tab:selected { color: #1a6b35; border-bottom: 3px solid #1e7145; font-weight: 600; }
        QTabBar::tab:hover:!selected { color: #1a6b35; background: #f0f9f4; }

        QComboBox {
            border: 1px solid #d0d0d0; border-radius: 3px;
            background: white; padding: 2px 6px;
            font-size: 11px; min-width: 40px;
        }
        QComboBox:hover { border-color: #1e7145; }
        QComboBox::drop-down { border: none; width: 18px; }
        QFontComboBox {
            border: 1px solid #d0d0d0; border-radius: 3px;
            background: white; font-size: 11px; min-width: 128px;
        }
        QFontComboBox:hover { border-color: #1e7145; }
        QSpinBox {
            border: 1px solid #d0d0d0; border-radius: 3px;
            background: white; font-size: 11px; min-width: 44px;
        }
        QSpinBox:hover { border-color: #1e7145; }
        QFrame[ribbon_sep="true"] { background: #e0e0e0; }
    )");

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(0,0,0,0);
    mainLayout->setSpacing(0);

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

    auto addSep   = [&]{ homeRow->addWidget(vSep()); };
    auto addSpace = [&](int px=4){ homeRow->addSpacing(px); };

    // ── CLIPBOARD ─────────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(0);

        auto* row = new QHBoxLayout; row->setSpacing(2);

        auto* btnPaste = makeLargeBtn(Ic::paste(),"Paste","Paste (Ctrl+V)",56,64);
        auto* vs = new QVBoxLayout; vs->setSpacing(1);
        auto* btnCut  = makeSmallBtn(Ic::cut(),     "Cut (Ctrl+X)");
        auto* btnCopy = makeSmallBtn(Ic::copy(),    "Copy (Ctrl+C)");
        auto* btnFmt  = makeSmallBtn(Ic::fmtpaint(),"Format Painter");
        vs->addWidget(btnCut); vs->addWidget(btnCopy); vs->addWidget(btnFmt);
        row->addWidget(btnPaste); row->addSpacing(2); row->addLayout(vs);
        col->addLayout(row); col->addWidget(groupLabel("Clipboard"));
        homeRow->addWidget(grp); addSep(); addSpace();

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

        auto* r1 = new QHBoxLayout; r1->setSpacing(2);
        d->fontFamily = new QFontComboBox;
        d->fontFamily->setCurrentFont(QFont("Calibri"));
        d->fontFamily->setFixedHeight(24); d->fontFamily->setMaximumWidth(142);
        d->fontSize = new QSpinBox;
        d->fontSize->setRange(6,96); d->fontSize->setValue(11);
        d->fontSize->setFixedHeight(24); d->fontSize->setFixedWidth(50);
        auto* btnGrow   = makeSmallBtn(makeIcon([](QPainter&p,int s){
            QFont f("Segoe UI",int(s*0.6),QFont::Bold); p.setFont(f); p.setPen(QColor("#333"));
            p.drawText(QRect(0,0,s-4,s),Qt::AlignCenter,"A");
            p.setPen(greenPen(1.5)); p.drawText(QRect(s/2,s/2-3,s/2+2,s/2+2),Qt::AlignCenter,"+");
        }),"Increase Font Size");
        auto* btnShrink = makeSmallBtn(makeIcon([](QPainter&p,int s){
            QFont f("Segoe UI",int(s*0.5),QFont::Bold); p.setFont(f); p.setPen(QColor("#333"));
            p.drawText(QRect(0,0,s-4,s),Qt::AlignCenter,"A");
            p.setPen(redPen(1.5)); p.drawText(QRect(s/2,s/2-3,s/2+2,s/2+2),Qt::AlignCenter,"-");
        }),"Decrease Font Size");
        r1->addWidget(d->fontFamily,1); r1->addWidget(d->fontSize);
        r1->addWidget(btnGrow); r1->addWidget(btnShrink);

        auto* r2 = new QHBoxLayout; r2->setSpacing(1);
        d->btnBold      = makeSmallBtn(Ic::bold(),      "Bold (Ctrl+B)",      true);
        d->btnItalic    = makeSmallBtn(Ic::italic(),    "Italic (Ctrl+I)",    true);
        d->btnUnderline = makeSmallBtn(Ic::underline(), "Underline (Ctrl+U)", true);
        d->btnBorders   = makeSmallBtn(Ic::borders(),   "Borders");
        d->btnTextColor = new ColorButton("", grp);
        d->btnTextColor->setToolTip("Font Color");     d->btnTextColor->setFixedSize(28,28);
        d->btnFillColor = new ColorButton("", grp);
        d->btnFillColor->setColor(Qt::yellow);
        d->btnFillColor->setToolTip("Highlight Color"); d->btnFillColor->setFixedSize(28,28);
        r2->addWidget(d->btnBold); r2->addWidget(d->btnItalic);
        r2->addWidget(d->btnUnderline); r2->addSpacing(3);
        r2->addWidget(d->btnBorders); r2->addSpacing(3);
        r2->addWidget(d->btnTextColor); r2->addWidget(d->btnFillColor);
        r2->addStretch();

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Font"));
        homeRow->addWidget(grp); addSep(); addSpace();

        connect(d->fontFamily,&QFontComboBox::currentFontChanged,this,
                [this](const QFont&f){ emit fontFamilyChanged(f.family()); });
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

        auto* r1 = new QHBoxLayout; r1->setSpacing(1);
        d->btnVTop = makeSmallBtn(Ic::vTop(),"Align Top",   true);
        d->btnVMid = makeSmallBtn(Ic::vMid(),"Align Middle",true);
        d->btnVBot = makeSmallBtn(Ic::vBot(),"Align Bottom",true);
        // Orientation dropdown
        auto* btnOri = makeSmallBtn(Ic::orientation(),"Orientation");
        auto* oriMenu = new QMenu(btnOri);
        oriMenu->addAction("Angle Counterclockwise");
        oriMenu->addAction("Angle Clockwise");
        oriMenu->addAction("Vertical Text");
        oriMenu->addAction("Rotate Text Up");
        oriMenu->addAction("Rotate Text Down");
        btnOri->setMenu(oriMenu); btnOri->setPopupMode(QToolButton::InstantPopup);
        d->btnWrap  = makeSmallBtn(Ic::wrap(),"Wrap Text",true);
        // Merge dropdown
        d->btnMerge = makeSmallBtn(Ic::merge(),"Merge and Center");
        auto* mergeMenu = new QMenu(d->btnMerge);
        mergeMenu->addAction("Merge and Center");
        mergeMenu->addAction("Merge Across");
        mergeMenu->addAction("Merge Cells");
        mergeMenu->addAction("Unmerge Cells");
        d->btnMerge->setMenu(mergeMenu); d->btnMerge->setPopupMode(QToolButton::MenuButtonPopup);

        r1->addWidget(d->btnVTop); r1->addWidget(d->btnVMid); r1->addWidget(d->btnVBot);
        r1->addSpacing(2); r1->addWidget(btnOri); r1->addWidget(d->btnWrap);
        r1->addSpacing(2); r1->addWidget(d->btnMerge); r1->addStretch();

        auto* r2 = new QHBoxLayout; r2->setSpacing(1);
        d->btnAlignL = makeSmallBtn(Ic::alignL(),"Align Left",  true);
        d->btnAlignC = makeSmallBtn(Ic::alignC(),"Center",      true);
        d->btnAlignR = makeSmallBtn(Ic::alignR(),"Align Right", true);
        auto* btnIndL = makeSmallBtn(makeIcon([](QPainter&p,int){
            p.setPen(darkPen(1.4));
            p.drawLine(2,5,18,5); p.drawLine(6,9,18,9); p.drawLine(2,13,18,13); p.drawLine(6,17,18,17);
            p.setPen(greenPen(2)); QPolygon a; a<<QPoint(2,9)<<QPoint(2,17)<<QPoint(5,13); p.drawPolygon(a);
        }),"Increase Indent");
        auto* btnIndR = makeSmallBtn(makeIcon([](QPainter&p,int){
            p.setPen(darkPen(1.4));
            p.drawLine(2,5,18,5); p.drawLine(2,9,14,9); p.drawLine(2,13,18,13); p.drawLine(2,17,14,17);
            p.setPen(greenPen(2)); QPolygon a; a<<QPoint(18,9)<<QPoint(18,17)<<QPoint(15,13); p.drawPolygon(a);
        }),"Decrease Indent");
        r2->addWidget(d->btnAlignL); r2->addWidget(d->btnAlignC); r2->addWidget(d->btnAlignR);
        r2->addSpacing(3); r2->addWidget(btnIndL); r2->addWidget(btnIndR); r2->addStretch();

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Alignment"));
        homeRow->addWidget(grp); addSep(); addSpace();

        connect(d->btnAlignL,&QToolButton::clicked,this,[this]{ emit hAlignChanged(0); });
        connect(d->btnAlignC,&QToolButton::clicked,this,[this]{ emit hAlignChanged(1); });
        connect(d->btnAlignR,&QToolButton::clicked,this,[this]{ emit hAlignChanged(2); });
        connect(d->btnVTop,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(0); });
        connect(d->btnVMid,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(1); });
        connect(d->btnVBot,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(2); });
        connect(d->btnWrap,  &QToolButton::toggled,this,&RibbonWidget::wrapTextToggled);
        connect(d->btnMerge, &QToolButton::clicked,this,&RibbonWidget::mergeCellsRequested);
    }

    // ── NUMBER FORMAT ─────────────────────────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        // Format dropdown
        d->numFormat = new QComboBox;
        d->numFormat->addItems({"General","Number","Currency","Accounting",
                                "Short Date","Long Date","Time","Percentage",
                                "Fraction","Scientific","Text"});
        d->numFormat->setFixedHeight(24); d->numFormat->setFixedWidth(132);

        // Format conversion button (unique to WPS)
        auto* btnFmtConv = makeSmallBtn(Ic::fmtConvert(),"Format Conversion",false,92);
        btnFmtConv->setText(" Format Conversion");
        btnFmtConv->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
        btnFmtConv->setFixedSize(140,24);
        btnFmtConv->setStyleSheet(
            "QToolButton { border:1px solid #d0d0d0; border-radius:3px; "
            "  background:white; font-size:10px; color:#333; padding:0 4px; text-align:left; }"
            "QToolButton:hover { border-color:#1e7145; background:#f0f9f4; }"
        );
        auto* fmtMenu = new QMenu(btnFmtConv);
        fmtMenu->addAction("Convert to Number");
        fmtMenu->addAction("Convert to Text");
        fmtMenu->addAction("Convert to Date");
        fmtMenu->addAction("Remove Leading Zeros");
        btnFmtConv->setMenu(fmtMenu);
        btnFmtConv->setPopupMode(QToolButton::InstantPopup);

        auto* r2 = new QHBoxLayout; r2->setSpacing(1);
        auto* btnCurr = makeSmallBtn(Ic::currency(),"Currency ($)");
        auto* btnPct  = makeSmallBtn(Ic::percent(),"Percent Style (%)");
        auto* btnThou = makeSmallBtn(Ic::thousands(),"Thousands Separator");
        d->btnDecInc  = makeSmallBtn(Ic::decInc(),"Increase Decimal");
        d->btnDecDec  = makeSmallBtn(Ic::decDec(),"Decrease Decimal");
        r2->addWidget(btnCurr); r2->addWidget(btnPct); r2->addWidget(btnThou);
        r2->addSpacing(3); r2->addWidget(d->btnDecInc); r2->addWidget(d->btnDecDec);
        r2->addStretch();

        col->addWidget(d->numFormat);
        col->addWidget(btnFmtConv);
        col->addLayout(r2);
        col->addWidget(groupLabel("Number Format"));
        homeRow->addWidget(grp); addSep(); addSpace();

        connect(d->numFormat,qOverload<int>(&QComboBox::currentIndexChanged),this,&RibbonWidget::numberFormatChanged);
        connect(btnCurr,&QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(2); emit numberFormatChanged(2); });
        connect(btnPct, &QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(7); emit numberFormatChanged(7); });
        connect(d->btnDecInc,&QToolButton::clicked,this,&RibbonWidget::increaseDecimalRequested);
        connect(d->btnDecDec,&QToolButton::clicked,this,&RibbonWidget::decreaseDecimalRequested);
    }

    // ── STYLES (Conditional Formatting | Cell Styles | Table Styles) ──────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        auto* r1 = new QHBoxLayout; r1->setSpacing(2);
        auto* btnCond  = makeMedDropBtn(Ic::condFmt(),   "Conditional\nFormatting","Conditional Formatting",112,44);
        auto* condMenu = new QMenu(btnCond);
        condMenu->addAction("Highlight Cell Rules");
        condMenu->addAction("Top/Bottom Rules");
        condMenu->addAction("Data Bars");
        condMenu->addAction("Color Scales");
        condMenu->addAction("Icon Sets");
        condMenu->addSeparator();
        condMenu->addAction("Manage Rules...");
        btnCond->setMenu(condMenu);

        r1->addWidget(btnCond);

        auto* r2 = new QHBoxLayout; r2->setSpacing(2);
        auto* btnCellSt = makeMedDropBtn(Ic::cellStyle(), "Cell\nStyles","Cell Styles",90,44);
        auto* btnTabSt  = makeMedDropBtn(Ic::tableStyle(),"Table\nStyles","Table Styles",90,44);
        auto* cellMenu = new QMenu(btnCellSt);
        cellMenu->addAction("Good, Bad, Neutral");
        cellMenu->addAction("Data and Model");
        cellMenu->addAction("Titles and Headings");
        cellMenu->addAction("Themed Cell Styles");
        cellMenu->addAction("Number Format");
        btnCellSt->setMenu(cellMenu);
        auto* tabMenu = new QMenu(btnTabSt);
        tabMenu->addAction("Light Styles");
        tabMenu->addAction("Medium Styles");
        tabMenu->addAction("Dark Styles");
        btnTabSt->setMenu(tabMenu);
        r2->addWidget(btnCellSt); r2->addWidget(btnTabSt);

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Styles"));
        homeRow->addWidget(grp); addSep(); addSpace();
    }

    // ── CELLS (Rows & Columns | Sheets) ───────────────────────────────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        auto* r1 = new QHBoxLayout; r1->setSpacing(2);
        auto* btnInsRow = makeSmallBtn(Ic::insRow(),"Insert Row");
        auto* btnDelRow = makeSmallBtn(Ic::delRow(),"Delete Row");
        auto* btnFmtCel = makeSmallBtn(Ic::fmtCell(),"Format Cells");
        r1->addWidget(btnInsRow); r1->addWidget(btnDelRow); r1->addWidget(btnFmtCel);

        auto* r2 = new QHBoxLayout; r2->setSpacing(2);
        auto* btnRowCol = makeMedDropBtn(Ic::rowsAndCols(),"Rows and\nColumns","Row/Column Operations",112,40);
        auto* rowColMenu = new QMenu(btnRowCol);
        rowColMenu->addAction("Insert Rows");
        rowColMenu->addAction("Insert Columns");
        rowColMenu->addSeparator();
        rowColMenu->addAction("Delete Rows");
        rowColMenu->addAction("Delete Columns");
        rowColMenu->addSeparator();
        rowColMenu->addAction("Hide Rows");
        rowColMenu->addAction("Unhide Rows");
        rowColMenu->addAction("Row Height...");
        rowColMenu->addAction("Column Width...");
        rowColMenu->addAction("AutoFit Row Height");
        rowColMenu->addAction("AutoFit Column Width");
        btnRowCol->setMenu(rowColMenu);
        r2->addWidget(btnRowCol);

        auto* btnSheets = makeMedDropBtn(Ic::sheets(),"Sheets","Sheet Operations",90,40);
        auto* sheetMenu = new QMenu(btnSheets);
        sheetMenu->addAction("Insert Sheet...");
        sheetMenu->addAction("Delete Sheet");
        sheetMenu->addAction("Rename Sheet");
        sheetMenu->addAction("Move/Copy Sheet...");
        sheetMenu->addSeparator();
        sheetMenu->addAction("Hide Sheet");
        sheetMenu->addAction("Unhide Sheet...");
        sheetMenu->addAction("Protect Sheet...");
        btnSheets->setMenu(sheetMenu);
        r2->addWidget(btnSheets);

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Cells"));
        homeRow->addWidget(grp); addSep(); addSpace();

        connect(btnInsRow,&QToolButton::clicked,this,&RibbonWidget::insertRowRequested);
        connect(btnDelRow,&QToolButton::clicked,this,&RibbonWidget::deleteRowRequested);
        connect(btnFmtCel,&QToolButton::clicked,this,&RibbonWidget::formatCellsRequested);
    }

    // ── EDITING (AutoSum | Fill | Sort | Filter | Freeze | Find) ──────────
    {
        auto* grp = new QWidget;
        auto* col = new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(2);

        auto* r1 = new QHBoxLayout; r1->setSpacing(2);

        // AutoSum large button with dropdown
        auto* btnSum  = makeLargeBtn(Ic::autosum(),"AutoSum","AutoSum (Alt+=)",48,64);
        auto* sumMenu = new QMenu(btnSum);
        sumMenu->addAction("Sum");
        sumMenu->addAction("Average");
        sumMenu->addAction("Count Numbers");
        sumMenu->addAction("Max");
        sumMenu->addAction("Min");
        sumMenu->addSeparator();
        sumMenu->addAction("More Functions...");
        btnSum->setMenu(sumMenu); btnSum->setPopupMode(QToolButton::MenuButtonPopup);

        // Fill large button with dropdown
        auto* btnFill = makeLargeBtn(Ic::fill(),"Fill","Fill",40,64);
        auto* fillMenu = new QMenu(btnFill);
        fillMenu->addAction("Fill Down  (Ctrl+D)");
        fillMenu->addAction("Fill Right (Ctrl+R)");
        fillMenu->addAction("Fill Up");
        fillMenu->addAction("Fill Left");
        fillMenu->addSeparator();
        fillMenu->addAction("Series...");
        fillMenu->addAction("Flash Fill (Ctrl+E)");
        btnFill->setMenu(fillMenu); btnFill->setPopupMode(QToolButton::MenuButtonPopup);

        // Sort/Filter/Find/Freeze as vertical stack
        auto* vs = new QVBoxLayout; vs->setSpacing(1);

        auto* vr1 = new QHBoxLayout; vr1->setSpacing(1);
        auto* btnSortA  = makeSmallBtn(Ic::sortAsc(), "Sort A to Z");
        auto* btnSortD  = makeSmallBtn(Ic::sortDesc(),"Sort Z to A");
        // Sort dropdown
        auto* btnSort   = makeSmallBtn(makeIcon([](QPainter&p,int){
            p.setPen(darkPen(1.3));
            p.drawLine(2,5,18,5); p.drawLine(2,9,14,9); p.drawLine(2,13,18,13); p.drawLine(2,17,10,17);
            p.setPen(greenPen(1.8)); QPolygon a; a<<QPoint(14,14)<<QPoint(14,18)<<QPoint(18,16); p.drawPolygon(a);
        }), "Custom Sort...");
        vr1->addWidget(btnSortA); vr1->addWidget(btnSortD); vr1->addWidget(btnSort);

        auto* vr2 = new QHBoxLayout; vr2->setSpacing(1);
        auto* btnFilter = makeSmallBtn(Ic::filter(),    "Filter");
        auto* btnFreeze = makeSmallBtn(Ic::freezePanes(),"Freeze Panes");
        // Freeze dropdown
        auto* freezeMenu = new QMenu(btnFreeze);
        freezeMenu->addAction("Freeze Panes");
        freezeMenu->addAction("Freeze Top Row");
        freezeMenu->addAction("Freeze First Column");
        freezeMenu->addSeparator();
        freezeMenu->addAction("Unfreeze Panes");
        btnFreeze->setMenu(freezeMenu); btnFreeze->setPopupMode(QToolButton::MenuButtonPopup);
        auto* btnFind   = makeSmallBtn(Ic::findIcon(),  "Find & Replace (Ctrl+H)");
        vr2->addWidget(btnFilter); vr2->addWidget(btnFreeze); vr2->addWidget(btnFind);

        vs->addLayout(vr1); vs->addLayout(vr2);

        r1->addWidget(btnSum); r1->addWidget(btnFill); r1->addSpacing(3); r1->addLayout(vs);

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
    tabs->addTab(new QWidget, "Tools");

    mainLayout->addWidget(tabs);

    auto* bottomLine = new QFrame;
    bottomLine->setFixedHeight(1);
    bottomLine->setProperty("ribbon_sep", true);
    bottomLine->setStyleSheet("background: #e0e0e0;");
    mainLayout->addWidget(bottomLine);
}

RibbonWidget::~RibbonWidget() { delete d; }

void RibbonWidget::setFormatState(const RibbonFormatState& s) {
    QSignalBlocker b1(d->fontFamily), b2(d->fontSize),
                   b3(d->btnBold), b4(d->btnItalic),
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
