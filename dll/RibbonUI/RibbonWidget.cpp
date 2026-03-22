// ═══════════════════════════════════════════════════════════════════════════════
//  RibbonWidget.cpp — Modern polished ribbon matching HTML design
//  Green theme, proper group labels, crisp SVG icons, full dropdowns
// ═══════════════════════════════════════════════════════════════════════════════
#include "RibbonWidget.h"
#include "ColorButton.h"
#include <QTabWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QComboBox>
#include <QToolButton>
#include <QLabel>
#include <QFontComboBox>
#include <QSpinBox>
#include <QFrame>
#include <QSizePolicy>
#include <QPainter>
#include <QPainterPath>
#include <QPixmap>
#include <QPolygon>
#include <QMenu>
#include <QFont>
#include <functional>

// ════════════════════════════════════════════════════════════════════════════
//  ICON FACTORY — crisp 20×20 QPainter icons
// ════════════════════════════════════════════════════════════════════════════
static QIcon makeIcon(std::function<void(QPainter&,int)> draw, int sz=20) {
    QPixmap pm(sz,sz); pm.fill(Qt::transparent);
    QPainter p(&pm); p.setRenderHint(QPainter::Antialiasing); draw(p,sz);
    return QIcon(pm);
}
static QPen dp(qreal w=1.6){ return QPen(QColor("#444"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }
static QPen gp(qreal w=1.8){ return QPen(QColor("#1a7a45"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }
static QPen rp(qreal w=1.6){ return QPen(QColor("#c0392b"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }

namespace Ic {
static QIcon paste(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.2)); p.setBrush(QColor("#e8e8e8")); p.drawRoundedRect(3,6,14,12,1,1);
    p.setBrush(Qt::white); p.drawRoundedRect(6,2,8,6,1,1);
    p.setPen(gp(1.5)); p.drawRoundedRect(6,2,8,6,1,1);
    p.setPen(dp(1.0)); p.drawLine(5,11,14,11); p.drawLine(5,14,14,14);
});}
static QIcon cut(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.6));
    p.drawEllipse(1,12,6,6); p.drawEllipse(12,12,6,6);
    p.drawLine(4,12,10,6); p.drawLine(16,12,10,6);
    p.drawLine(4,2,10,6); p.drawLine(16,2,10,6);
});}
static QIcon copy(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e0e0e0")); p.drawRoundedRect(5,5,13,13,1,1);
    p.setBrush(Qt::white); p.drawRoundedRect(2,2,13,13,1,1);
    p.setBrush(Qt::NoBrush); p.setPen(dp(1.2)); p.drawRoundedRect(2,2,13,13,1,1);
});}
static QIcon fmtpaint(){ return makeIcon([](QPainter&p,int){
    p.setBrush(QColor("#1a7a45")); p.setPen(gp()); p.drawRoundedRect(1,1,9,12,1,1);
    p.setPen(dp(1.5)); p.setBrush(Qt::NoBrush);
    p.drawLine(10,11,16,17); p.drawLine(13,9,16,6);
});}
static QIcon bold(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.65),QFont::Black); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"B");
});}
static QIcon italic(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.65),QFont::Normal,true); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"I");
});}
static QIcon underline(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.55),QFont::Bold); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,-1,s,s),Qt::AlignCenter,"U");
    p.setPen(QPen(QColor("#1a7a45"),2)); p.drawLine(2,s-2,s-2,s-2);
});}
static QIcon borders(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3)); p.drawRect(2,2,16,16);
    p.setPen(QPen(QColor("#aaa"),0.7)); p.drawLine(10,2,10,18); p.drawLine(2,10,18,10);
});}
static QIcon alignL(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.6)); p.drawLine(2,5,18,5); p.drawLine(2,9,13,9); p.drawLine(2,13,18,13); p.drawLine(2,17,13,17);
});}
static QIcon alignC(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.6)); p.drawLine(2,5,18,5); p.drawLine(5,9,15,9); p.drawLine(2,13,18,13); p.drawLine(5,17,15,17);
});}
static QIcon alignR(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.6)); p.drawLine(2,5,18,5); p.drawLine(7,9,18,9); p.drawLine(2,13,18,13); p.drawLine(7,17,18,17);
});}
static QIcon vTop(){ return makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(2,2,18,2);
    p.setPen(dp(1.5)); p.drawLine(6,5,10,15); p.drawLine(14,5,10,15); p.drawLine(6,5,14,5);
});}
static QIcon vMid(){ return makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(2,10,18,10);
    p.setPen(dp(1.5)); p.drawRect(5,3,10,14);
});}
static QIcon vBot(){ return makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(2,18,18,18);
    p.setPen(dp(1.5)); p.drawLine(6,15,10,5); p.drawLine(14,15,10,5); p.drawLine(6,15,14,15);
});}
static QIcon wrap(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.4)); p.drawLine(2,5,16,5); p.drawLine(2,9,18,9);
    QPainterPath path; path.moveTo(18,9); path.cubicTo(20,9,20,14,18,14); path.lineTo(14,14);
    p.drawPath(path);
    p.setBrush(QColor("#444")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(14,11)<<QPoint(14,17)<<QPoint(10,14); p.drawPolygon(a);
    p.setPen(dp(1.4)); p.drawLine(2,13,10,13);
});}
static QIcon merge(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.1));
    p.drawRect(2,2,7,7); p.drawRect(11,2,7,7); p.drawRect(2,11,7,7); p.drawRect(11,11,7,7);
    p.setPen(gp(2)); p.drawLine(10,2,10,18); p.drawLine(2,10,18,10);
});}
static QIcon orient(){ return makeIcon([](QPainter&p,int){
    p.setPen(gp(1.8));
    p.save(); p.translate(10,10); p.rotate(-45); p.drawLine(-7,0,7,0);
    p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(4,-3)<<QPoint(7,0)<<QPoint(4,3); p.drawPolygon(a);
    p.restore();
    p.setPen(dp(1.3)); p.drawLine(2,17,18,17);
});}
static QIcon currency(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.65),QFont::Bold); p.setFont(f); p.setPen(QColor("#1a7a45"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"$");
});}
static QIcon percent(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.65),QFont::Bold); p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"%");
});}
static QIcon thousands(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.5)); p.setFont(f); p.setPen(QColor("#444"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,",");
});}
static QIcon decInc(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.42)); p.setFont(f); p.setPen(gp(1.4));
    p.drawText(QRect(0,0,s/2+2,s),Qt::AlignVCenter|Qt::AlignRight,".0");
    p.setPen(QColor("#444")); p.drawText(QRect(s/2,0,s/2,s),Qt::AlignVCenter|Qt::AlignLeft,">");
});}
static QIcon decDec(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.42)); p.setFont(f); p.setPen(rp(1.4));
    p.drawText(QRect(0,0,s/2+4,s),Qt::AlignVCenter|Qt::AlignRight,"<");
    p.setPen(QColor("#444")); p.drawText(QRect(s/2,0,s/2,s),Qt::AlignVCenter|Qt::AlignLeft,".0");
});}
static QIcon condFmt(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.1)); p.setBrush(QColor("#ffe8e8")); p.drawRoundedRect(2,2,16,7,1,1);
    p.setBrush(QColor("#e8ffe8")); p.drawRoundedRect(2,11,16,7,1,1);
    p.setPen(rp(1.3)); p.drawLine(4,5,7,5); p.drawLine(9,5,14,5);
    p.setPen(gp(1.3)); p.drawLine(4,14,14,14);
});}
static QIcon cellStyle(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.0)); p.setBrush(QColor("#c8e8d8")); p.drawRect(2,2,7,7);
    p.setBrush(QColor("#b8d4f0")); p.drawRect(11,2,7,7);
    p.setBrush(QColor("#ffd8a8")); p.drawRect(2,11,7,7);
    p.setBrush(QColor("#f0c8c8")); p.drawRect(11,11,7,7);
});}
static QIcon tableStyle(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#1a7a45")); p.drawRect(2,2,16,5);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,7,8,4); p.setBrush(Qt::white); p.drawRect(10,7,8,4);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,11,8,4); p.setBrush(Qt::white); p.drawRect(10,11,8,4);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,15,8,3); p.setBrush(Qt::white); p.drawRect(10,15,8,3);
});}
static QIcon rowsCols(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3));
    p.drawLine(2,5,18,5); p.drawLine(2,9,18,9); p.drawLine(2,13,18,13); p.drawLine(2,17,18,17);
    p.setPen(gp(1.8)); p.drawLine(10,2,10,18);
    p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
    QPolygon u; u<<QPoint(7,3)<<QPoint(10,0)<<QPoint(13,3); p.drawPolygon(u);
    QPolygon d; d<<QPoint(7,17)<<QPoint(10,20)<<QPoint(13,17); p.drawPolygon(d);
});}
static QIcon sheets(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e8e8e8")); p.drawRoundedRect(2,6,18,13,1,1);
    p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(4,4,8,5,1,1);
    p.setPen(gp(1.5)); p.drawLine(14,3,14,8); p.drawLine(11,6,17,6);
});}
static QIcon autosum(){ return makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",int(s*0.72),QFont::Bold); p.setFont(f); p.setPen(QColor("#1a7a45"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"Σ");
});}
static QIcon fill(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(2,2,16,16,2,2);
    p.setPen(gp(2.2)); p.drawLine(10,5,10,15); p.drawLine(5,10,15,10);
});}
static QIcon sortAsc(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.5));
    p.drawLine(3,16,8,4); p.drawLine(8,4,13,16); p.drawLine(5,12,11,12);
    p.setPen(dp(1.2)); p.drawLine(15,5,19,5); p.drawLine(15,9,19,9); p.drawLine(15,13,19,13);
});}
static QIcon sortDesc(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.5));
    p.drawLine(3,4,8,16); p.drawLine(8,16,13,4); p.drawLine(5,8,11,8);
    p.setPen(dp(1.2)); p.drawLine(15,5,19,5); p.drawLine(15,9,19,9); p.drawLine(15,13,19,13);
});}
static QIcon filter(){ return makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.setBrush(Qt::NoBrush);
    p.drawLine(2,4,18,4); p.drawLine(5,9,15,9); p.drawLine(8,14,12,14);
    p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(5,18)<<QPoint(8,15)<<QPoint(8,9)<<QPoint(5,9); p.drawPolygon(a);
});}
static QIcon findIcon(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.8)); p.setBrush(Qt::NoBrush);
    p.drawEllipse(2,2,12,12); p.drawLine(11,11,18,18);
});}
static QIcon freeze(){ return makeIcon([](QPainter&p,int){
    p.setPen(gp(2.2)); p.drawLine(10,2,10,18); p.drawLine(2,8,18,8);
    p.setPen(dp(1.0));
    p.drawRect(2,2,8,6); p.drawRect(2,8,8,10); p.drawRect(10,2,8,6); p.drawRect(10,8,8,10);
});}
static QIcon insRow(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.2)); p.drawLine(2,6,18,6); p.drawLine(2,14,18,14);
    p.setPen(gp(2.2)); p.drawLine(10,8,10,12); p.drawLine(7,10,13,10);
});}
static QIcon delRow(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp(1.2)); p.drawLine(2,6,18,6); p.drawLine(2,14,18,14);
    p.setPen(rp(2.2)); p.drawLine(7,10,13,10);
});}
static QIcon fmtCell(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(Qt::white); p.drawRoundedRect(2,2,16,16,2,2);
    p.setPen(gp(1.5)); p.drawLine(5,7,15,7); p.drawLine(5,11,15,11); p.drawLine(5,15,11,15);
});}
static QIcon fmtConv(){ return makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(1,2,11,16,2,2);
    p.setPen(gp(2)); p.drawLine(13,10,19,10);
    p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(16,7)<<QPoint(19,10)<<QPoint(16,13); p.drawPolygon(a);
});}
} // namespace Ic

// ════════════════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════════════════
static QFrame* vSep(){
    auto*f=new QFrame; f->setFrameShape(QFrame::VLine);
    f->setFixedWidth(1); f->setFixedHeight(74); f->setStyleSheet("color:#e0e3e8;"); return f;
}

static QLabel* groupLabel(const QString& text){
    auto*l=new QLabel(text); l->setAlignment(Qt::AlignCenter);
    l->setStyleSheet("color:#999;font-size:10px;font-family:'Segoe UI',Arial;font-weight:500;");
    l->setFixedHeight(14); return l;
}

// Large icon-label button (Paste, AutoSum, Fill)
static QToolButton* makeLargeBtn(const QIcon& icon, const QString& label,
                                   const QString& tip, int w=54, int h=68){
    auto*btn=new QToolButton;
    btn->setIcon(icon); btn->setIconSize(QSize(28,28));
    btn->setText(label); btn->setToolTip(tip);
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    btn->setAutoRaise(true); btn->setFixedSize(w,h);
    btn->setStyleSheet(
        "QToolButton{border:1px solid transparent;border-radius:4px;"
        "background:transparent;font-size:10px;color:#333;padding:3px 2px 2px;font-family:'Segoe UI';}"
        "QToolButton:hover{background:#e8f5ee;border-color:#b8ddc8;}"
        "QToolButton:pressed{background:#c8e8d8;border-color:#1a7a45;}");
    return btn;
}

// Small square button (28×28)
static QToolButton* makeSmBtn(const QIcon& icon, const QString& tip,
                               bool checkable=false, int sz=28){
    auto*btn=new QToolButton;
    btn->setIcon(icon); btn->setIconSize(QSize(sz-8,sz-8));
    btn->setToolTip(tip); btn->setAutoRaise(true);
    btn->setCheckable(checkable); btn->setFixedSize(sz,sz);
    btn->setStyleSheet(
        "QToolButton{border:1px solid transparent;border-radius:3px;background:transparent;}"
        "QToolButton:hover{background:#e8f5ee;border-color:#b8ddc8;}"
        "QToolButton:pressed,QToolButton:checked{background:#c8e8d8;border-color:#1a7a45;}");
    return btn;
}

// Medium drop button with icon + 2-line label
static QToolButton* makeMedBtn(const QIcon& icon, const QString& label,
                                const QString& tip, int w=110, int h=44){
    auto*btn=new QToolButton;
    btn->setIcon(icon); btn->setIconSize(QSize(22,22));
    btn->setText(label); btn->setToolTip(tip);
    btn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    btn->setPopupMode(QToolButton::MenuButtonPopup);
    btn->setAutoRaise(true); btn->setFixedSize(w,h);
    btn->setStyleSheet(
        "QToolButton{border:1px solid transparent;border-radius:3px;"
        "background:transparent;font-size:10px;color:#333;padding:2px 4px;font-family:'Segoe UI';}"
        "QToolButton:hover{background:#e8f5ee;border-color:#b8ddc8;}"
        "QToolButton:pressed{background:#c8e8d8;border-color:#1a7a45;}"
        "QToolButton::menu-button{width:14px;border:none;}");
    return btn;
}

// Styled number symbol button ($, %, ,)
static QToolButton* makeSymBtn(const QString& sym, const QString& tip){
    auto*btn=new QToolButton;
    btn->setText(sym); btn->setToolTip(tip); btn->setAutoRaise(true);
    btn->setFixedSize(26,26);
    btn->setStyleSheet(
        "QToolButton{border:1px solid #dde1e7;border-radius:3px;background:white;"
        "font-size:13px;font-weight:600;color:#1a7a45;font-family:'Segoe UI';}"
        "QToolButton:hover{background:#e8f5ee;border-color:#1a7a45;}"
        "QToolButton:pressed{background:#c8e8d8;}");
    return btn;
}

static QMenu* simpleMenu(QWidget* parent, QStringList items){
    auto*m=new QMenu(parent);
    m->setStyleSheet(
        "QMenu{background:white;border:1px solid #dde1e7;border-radius:5px;padding:4px 0;font-family:'Segoe UI';font-size:12px;}"
        "QMenu::item{padding:7px 20px 7px 14px;color:#1a1d23;}"
        "QMenu::item:selected{background:#e8f5ee;color:#1a7a45;}"
        "QMenu::separator{height:1px;background:#dde1e7;margin:3px 0;}");
    for(const auto&s:items){
        if(s=="---") m->addSeparator();
        else m->addAction(s);
    }
    return m;
}

// ════════════════════════════════════════════════════════════════════════════
//  Impl
// ════════════════════════════════════════════════════════════════════════════
struct RibbonWidget::Impl {
    QFontComboBox* fontFamily{nullptr};
    QSpinBox*      fontSize  {nullptr};
    QToolButton   *btnBold{nullptr},*btnItalic{nullptr},*btnUnderline{nullptr};
    QToolButton   *btnBorders{nullptr};
    ColorButton   *btnTextColor{nullptr},*btnFillColor{nullptr};
    QToolButton   *btnAlignL{nullptr},*btnAlignC{nullptr},*btnAlignR{nullptr};
    QToolButton   *btnVTop{nullptr},*btnVMid{nullptr},*btnVBot{nullptr};
    QToolButton   *btnWrap{nullptr},*btnMerge{nullptr};
    QComboBox     *numFormat{nullptr};
    QToolButton   *btnDecInc{nullptr},*btnDecDec{nullptr};
};

// ════════════════════════════════════════════════════════════════════════════
//  CONSTRUCTION
// ════════════════════════════════════════════════════════════════════════════
RibbonWidget::RibbonWidget(QWidget* parent)
    : QWidget(parent), d(new Impl)
{
    setFixedHeight(136);
    setAttribute(Qt::WA_StyledBackground, true);
    setStyleSheet(
        "RibbonWidget{background:#ffffff;border-bottom:1px solid #dde1e7;}"
        "QTabWidget::pane{border:none;background:#ffffff;margin-top:-1px;}"
        "QTabBar::tab{"
        "  padding:5px 16px 4px;font-size:12px;font-weight:500;"
        "  color:#666;background:transparent;border:none;"
        "  border-bottom:2px solid transparent;min-width:46px;"
        "  font-family:'Segoe UI',Arial;"
        "}"
        "QTabBar::tab:selected{color:#1a7a45;border-bottom:2px solid #1a7a45;font-weight:600;}"
        "QTabBar::tab:hover:!selected{color:#1a7a45;background:#f0f9f4;}"
        "QComboBox{"
        "  border:1px solid #dde1e7;border-radius:3px;background:white;"
        "  padding:2px 6px;font-size:12px;font-family:'Segoe UI',Arial;"
        "  selection-background-color:#c8e8d8;"
        "}"
        "QComboBox:hover,QComboBox:focus{border-color:#1a7a45;}"
        "QComboBox::drop-down{border:none;width:18px;}"
        "QFontComboBox{"
        "  border:1px solid #dde1e7;border-radius:3px;background:white;"
        "  font-size:12px;min-width:130px;font-family:'Segoe UI',Arial;"
        "}"
        "QFontComboBox:hover{border-color:#1a7a45;}"
        "QSpinBox{"
        "  border:1px solid #dde1e7;border-radius:3px;background:white;"
        "  font-size:12px;min-width:44px;font-family:'Segoe UI',Arial;"
        "}"
        "QSpinBox:hover{border-color:#1a7a45;}"
    );

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(0,0,0,0);
    mainLayout->setSpacing(0);

    // Tab bar + ribbon body
    auto* tabs = new QTabWidget;
    tabs->setTabPosition(QTabWidget::North);
    tabs->setDocumentMode(true);
    tabs->setElideMode(Qt::ElideNone);

    // ════════════════════ HOME TAB ════════════════════════════════════════
    auto* homeTab = new QWidget;
    homeTab->setStyleSheet("background:#ffffff;");
    auto* homeRow = new QHBoxLayout(homeTab);
    homeRow->setContentsMargins(8,2,8,0);
    homeRow->setSpacing(0);

    auto addSep=[&]{ homeRow->addWidget(vSep()); };
    auto addSp =[&](int px=4){ homeRow->addSpacing(px); };

    // ── CLIPBOARD ────────────────────────────────────────────────────────
    {
        auto*grp=new QWidget; auto*col=new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(0);
        auto*row=new QHBoxLayout; row->setSpacing(2);

        auto*btnPaste=makeLargeBtn(Ic::paste(),"Paste","Paste (Ctrl+V)",56,70);
        auto*vs=new QVBoxLayout; vs->setSpacing(2);
        auto*btnCut =makeSmBtn(Ic::cut(),  "Cut (Ctrl+X)");
        auto*btnCopy=makeSmBtn(Ic::copy(), "Copy (Ctrl+C)");
        auto*btnFmt =makeSmBtn(Ic::fmtpaint(),"Format Painter");
        vs->addWidget(btnCut); vs->addWidget(btnCopy); vs->addWidget(btnFmt);
        row->addWidget(btnPaste); row->addSpacing(3); row->addLayout(vs);
        col->addLayout(row); col->addWidget(groupLabel("Clipboard"));
        homeRow->addWidget(grp); addSep(); addSp();

        connect(btnPaste,&QToolButton::clicked,this,&RibbonWidget::pasteRequested);
        connect(btnCut,  &QToolButton::clicked,this,&RibbonWidget::cutRequested);
        connect(btnCopy, &QToolButton::clicked,this,&RibbonWidget::copyRequested);
        connect(btnFmt,  &QToolButton::clicked,this,&RibbonWidget::formatPainterRequested);
    }

    // ── FONT ─────────────────────────────────────────────────────────────
    {
        auto*grp=new QWidget; auto*col=new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(3);

        // Row 1: font + size + grow/shrink
        auto*r1=new QHBoxLayout; r1->setSpacing(3);
        d->fontFamily=new QFontComboBox;
        d->fontFamily->setCurrentFont(QFont("Calibri"));
        d->fontFamily->setFixedHeight(26); d->fontFamily->setMaximumWidth(148);
        d->fontSize=new QSpinBox;
        d->fontSize->setRange(6,96); d->fontSize->setValue(11);
        d->fontSize->setFixedHeight(26); d->fontSize->setFixedWidth(50);

        auto*btnGrow  =makeSmBtn(makeIcon([](QPainter&p,int s){
            QFont f("Segoe UI",int(s*0.55),QFont::Bold); p.setFont(f); p.setPen(QColor("#333"));
            p.drawText(QRect(0,0,s-4,s),Qt::AlignCenter,"A");
            p.setPen(gp(1.5)); p.drawText(QRect(s/2-1,s/2-4,s/2+2,s/2+3),Qt::AlignCenter,"+");
        }),"Increase Font Size");
        auto*btnShrink=makeSmBtn(makeIcon([](QPainter&p,int s){
            QFont f("Segoe UI",int(s*0.48)); p.setFont(f); p.setPen(QColor("#333"));
            p.drawText(QRect(0,0,s-4,s),Qt::AlignCenter,"A");
            p.setPen(rp(1.5)); p.drawText(QRect(s/2-1,s/2-4,s/2+2,s/2+3),Qt::AlignCenter,"-");
        }),"Decrease Font Size");

        r1->addWidget(d->fontFamily,1); r1->addWidget(d->fontSize);
        r1->addWidget(btnGrow); r1->addWidget(btnShrink);

        // Row 2: B I U | border | text-color fill-color
        auto*r2=new QHBoxLayout; r2->setSpacing(2);
        d->btnBold     =makeSmBtn(Ic::bold(),     "Bold (Ctrl+B)",      true);
        d->btnItalic   =makeSmBtn(Ic::italic(),   "Italic (Ctrl+I)",    true);
        d->btnUnderline=makeSmBtn(Ic::underline(),"Underline (Ctrl+U)", true);
        d->btnBorders  =makeSmBtn(Ic::borders(),  "Borders");
        d->btnTextColor=new ColorButton("",grp);
        d->btnTextColor->setToolTip("Font Color"); d->btnTextColor->setFixedSize(28,28);
        d->btnFillColor=new ColorButton("",grp);
        d->btnFillColor->setColor(Qt::yellow);
        d->btnFillColor->setToolTip("Highlight Color"); d->btnFillColor->setFixedSize(28,28);

        r2->addWidget(d->btnBold); r2->addWidget(d->btnItalic);
        r2->addWidget(d->btnUnderline); r2->addSpacing(4);
        r2->addWidget(d->btnBorders); r2->addSpacing(4);
        r2->addWidget(d->btnTextColor); r2->addWidget(d->btnFillColor);
        r2->addStretch();

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Font"));
        homeRow->addWidget(grp); addSep(); addSp();

        connect(d->fontFamily,&QFontComboBox::currentFontChanged,this,
            [this](const QFont&f){ emit fontFamilyChanged(f.family()); });
        connect(d->fontSize,qOverload<int>(&QSpinBox::valueChanged),this,&RibbonWidget::fontSizeChanged);
        connect(d->btnBold,      &QToolButton::toggled,this,&RibbonWidget::boldToggled);
        connect(d->btnItalic,    &QToolButton::toggled,this,&RibbonWidget::italicToggled);
        connect(d->btnUnderline, &QToolButton::toggled,this,&RibbonWidget::underlineToggled);
        connect(d->btnTextColor, &ColorButton::colorChanged,this,&RibbonWidget::textColorChanged);
        connect(d->btnFillColor, &ColorButton::colorChanged,this,&RibbonWidget::fillColorChanged);
    }

    // ── ALIGNMENT ────────────────────────────────────────────────────────
    {
        auto*grp=new QWidget; auto*col=new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(3);

        auto*r1=new QHBoxLayout; r1->setSpacing(2);
        d->btnVTop=makeSmBtn(Ic::vTop(),"Align Top",   true);
        d->btnVMid=makeSmBtn(Ic::vMid(),"Align Middle",true);
        d->btnVBot=makeSmBtn(Ic::vBot(),"Align Bottom",true);

        auto*btnOri=makeSmBtn(Ic::orient(),"Orientation");
        auto*oriMenu=simpleMenu(btnOri,{"Angle Counterclockwise","Angle Clockwise","Vertical Text","---","Rotate Text Up","Rotate Text Down"});
        btnOri->setMenu(oriMenu); btnOri->setPopupMode(QToolButton::InstantPopup);

        d->btnWrap =makeSmBtn(Ic::wrap(), "Wrap Text",true);
        d->btnMerge=makeSmBtn(Ic::merge(),"Merge Cells");
        auto*mergeMenu=simpleMenu(d->btnMerge,{"Merge and Center","Merge Across","---","Unmerge Cells"});
        d->btnMerge->setMenu(mergeMenu); d->btnMerge->setPopupMode(QToolButton::MenuButtonPopup);

        r1->addWidget(d->btnVTop); r1->addWidget(d->btnVMid); r1->addWidget(d->btnVBot);
        r1->addSpacing(3); r1->addWidget(btnOri); r1->addWidget(d->btnWrap);
        r1->addSpacing(2); r1->addWidget(d->btnMerge); r1->addStretch();

        auto*r2=new QHBoxLayout; r2->setSpacing(2);
        d->btnAlignL=makeSmBtn(Ic::alignL(),"Align Left",  true);
        d->btnAlignC=makeSmBtn(Ic::alignC(),"Center",      true);
        d->btnAlignR=makeSmBtn(Ic::alignR(),"Align Right", true);

        auto*btnIndDec=makeSmBtn(makeIcon([](QPainter&p,int){
            p.setPen(dp(1.4)); p.drawLine(2,5,18,5); p.drawLine(6,9,18,9); p.drawLine(2,13,18,13); p.drawLine(6,17,18,17);
            p.setPen(gp(2)); QPolygon a; a<<QPoint(2,9)<<QPoint(2,17)<<QPoint(5,13); p.drawPolygon(a);
        }),"Increase Indent");
        auto*btnIndInc=makeSmBtn(makeIcon([](QPainter&p,int){
            p.setPen(dp(1.4)); p.drawLine(2,5,18,5); p.drawLine(2,9,14,9); p.drawLine(2,13,18,13); p.drawLine(2,17,14,17);
            p.setPen(gp(2)); QPolygon a; a<<QPoint(18,9)<<QPoint(18,17)<<QPoint(15,13); p.drawPolygon(a);
        }),"Decrease Indent");

        r2->addWidget(d->btnAlignL); r2->addWidget(d->btnAlignC); r2->addWidget(d->btnAlignR);
        r2->addSpacing(3); r2->addWidget(btnIndDec); r2->addWidget(btnIndInc); r2->addStretch();

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Alignment"));
        homeRow->addWidget(grp); addSep(); addSp();

        connect(d->btnAlignL,&QToolButton::clicked,this,[this]{ emit hAlignChanged(0); });
        connect(d->btnAlignC,&QToolButton::clicked,this,[this]{ emit hAlignChanged(1); });
        connect(d->btnAlignR,&QToolButton::clicked,this,[this]{ emit hAlignChanged(2); });
        connect(d->btnVTop,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(0); });
        connect(d->btnVMid,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(1); });
        connect(d->btnVBot,  &QToolButton::clicked,this,[this]{ emit vAlignChanged(2); });
        connect(d->btnWrap,  &QToolButton::toggled,this,&RibbonWidget::wrapTextToggled);
        connect(d->btnMerge, &QToolButton::clicked,this,&RibbonWidget::mergeCellsRequested);
    }

    // ── NUMBER FORMAT ────────────────────────────────────────────────────
    {
        auto*grp=new QWidget; auto*col=new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(3);

        d->numFormat=new QComboBox;
        d->numFormat->addItems({"General","Number","Currency ($)","Accounting",
                                "Short Date","Long Date","Time","Percentage (%)",
                                "Fraction","Scientific","Text"});
        d->numFormat->setFixedHeight(26); d->numFormat->setFixedWidth(148);

        // Format Conversion button (WPS exclusive)
        auto*btnFmtConv=new QToolButton;
        btnFmtConv->setIcon(Ic::fmtConv()); btnFmtConv->setIconSize(QSize(16,16));
        btnFmtConv->setText(" Format Conversion");
        btnFmtConv->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
        btnFmtConv->setFixedSize(148,26); btnFmtConv->setPopupMode(QToolButton::InstantPopup);
        btnFmtConv->setStyleSheet(
            "QToolButton{border:1px solid #dde1e7;border-radius:3px;background:#f7f8fa;"
            "font-size:11px;color:#1a7a45;font-weight:500;font-family:'Segoe UI';padding:0 6px;text-align:left;}"
            "QToolButton:hover{border-color:#1a7a45;background:#e8f5ee;}"
            "QToolButton::menu-button{width:14px;border:none;}");
        btnFmtConv->setMenu(simpleMenu(btnFmtConv,{"Convert to Number","Convert to Text","Convert to Date","---","Remove Leading Zeros"}));

        auto*r2=new QHBoxLayout; r2->setSpacing(2);
        auto*btnCurr=makeSymBtn("$","Currency ($)");
        auto*btnPct =makeSymBtn("%","Percentage (%)");
        auto*btnThou=makeSymBtn(",","Thousands Separator");
        d->btnDecInc=makeSmBtn(Ic::decInc(),"Increase Decimal");
        d->btnDecDec=makeSmBtn(Ic::decDec(),"Decrease Decimal");
        r2->addWidget(btnCurr); r2->addWidget(btnPct); r2->addWidget(btnThou);
        r2->addSpacing(4); r2->addWidget(d->btnDecInc); r2->addWidget(d->btnDecDec);
        r2->addStretch();

        col->addWidget(d->numFormat);
        col->addWidget(btnFmtConv);
        col->addLayout(r2);
        col->addWidget(groupLabel("Number Format"));
        homeRow->addWidget(grp); addSep(); addSp();

        connect(d->numFormat,qOverload<int>(&QComboBox::currentIndexChanged),this,&RibbonWidget::numberFormatChanged);
        connect(btnCurr,&QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(2); emit numberFormatChanged(2); });
        connect(btnPct, &QToolButton::clicked,this,[this]{ d->numFormat->setCurrentIndex(7); emit numberFormatChanged(7); });
        connect(d->btnDecInc,&QToolButton::clicked,this,&RibbonWidget::increaseDecimalRequested);
        connect(d->btnDecDec,&QToolButton::clicked,this,&RibbonWidget::decreaseDecimalRequested);
    }

    // ── STYLES ────────────────────────────────────────────────────────────
    {
        auto*grp=new QWidget; auto*col=new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(3);

        auto*btnCond=makeMedBtn(Ic::condFmt(),"Conditional\nFormatting","Conditional Formatting",136,40);
        btnCond->setMenu(simpleMenu(btnCond,{"Highlight Cell Rules","Top/Bottom Rules","---","Data Bars","Color Scales","Icon Sets","---","Manage Rules..."}));

        auto*r2=new QHBoxLayout; r2->setSpacing(3);
        auto*btnCellSt=makeMedBtn(Ic::cellStyle(),"Cell\nStyles","Cell Styles",88,36);
        btnCellSt->setMenu(simpleMenu(btnCellSt,{"Good","Bad","Neutral","---","Heading 1","Heading 2","Total","---","New Cell Style..."}));
        auto*btnTabSt=makeMedBtn(Ic::tableStyle(),"Table\nStyles","Table Styles",88,36);
        btnTabSt->setMenu(simpleMenu(btnTabSt,{"Light 1 (Green)","Light 2 (Blue)","Medium 1","Dark 1"}));
        r2->addWidget(btnCellSt); r2->addWidget(btnTabSt);

        col->addWidget(btnCond); col->addLayout(r2);
        col->addWidget(groupLabel("Styles"));
        homeRow->addWidget(grp); addSep(); addSp();
    }

    // ── CELLS ─────────────────────────────────────────────────────────────
    {
        auto*grp=new QWidget; auto*col=new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(3);

        auto*r1=new QHBoxLayout; r1->setSpacing(2);
        auto*btnInsRow=makeSmBtn(Ic::insRow(),"Insert Row");
        auto*btnDelRow=makeSmBtn(Ic::delRow(),"Delete Row");
        auto*btnFmtCel=makeSmBtn(Ic::fmtCell(),"Format Cells");
        r1->addWidget(btnInsRow); r1->addWidget(btnDelRow); r1->addWidget(btnFmtCel);

        auto*btnRowCol=makeMedBtn(Ic::rowsCols(),"Rows and\nColumns","Row/Column Operations",120,40);
        btnRowCol->setMenu(simpleMenu(btnRowCol,{"Insert Rows","Insert Columns","---","Delete Rows","Delete Columns","---","Hide Rows","Unhide Rows","---","Row Height...","Column Width...","AutoFit Row Height","AutoFit Column Width"}));

        auto*btnSheets=makeMedBtn(Ic::sheets(),"Sheets","Sheet Operations",100,40);
        btnSheets->setMenu(simpleMenu(btnSheets,{"Insert Sheet...","Delete Sheet","Rename Sheet","Move/Copy Sheet...","---","Hide Sheet","Unhide Sheet...","---","Protect Sheet..."}));

        auto*r2=new QHBoxLayout; r2->setSpacing(3);
        r2->addWidget(btnRowCol); r2->addWidget(btnSheets);

        col->addLayout(r1); col->addLayout(r2);
        col->addWidget(groupLabel("Cells"));
        homeRow->addWidget(grp); addSep(); addSp();

        connect(btnInsRow,&QToolButton::clicked,this,&RibbonWidget::insertRowRequested);
        connect(btnDelRow,&QToolButton::clicked,this,&RibbonWidget::deleteRowRequested);
        connect(btnFmtCel,&QToolButton::clicked,this,&RibbonWidget::formatCellsRequested);
    }

    // ── EDITING ───────────────────────────────────────────────────────────
    {
        auto*grp=new QWidget; auto*col=new QVBoxLayout(grp);
        col->setContentsMargins(2,0,2,0); col->setSpacing(0);

        auto*r1=new QHBoxLayout; r1->setSpacing(2);

        auto*btnSum=makeLargeBtn(Ic::autosum(),"AutoSum","AutoSum (Alt+=)",52,70);
        btnSum->setMenu(simpleMenu(btnSum,{"Sum","Average","Count Numbers","Max","Min","---","More Functions..."}));
        btnSum->setPopupMode(QToolButton::MenuButtonPopup);

        auto*btnFill=makeLargeBtn(Ic::fill(),"Fill","Fill",44,70);
        btnFill->setMenu(simpleMenu(btnFill,{"Fill Down  Ctrl+D","Fill Right  Ctrl+R","Fill Up","Fill Left","---","Series...","Flash Fill  Ctrl+E"}));
        btnFill->setPopupMode(QToolButton::MenuButtonPopup);

        auto*vs=new QVBoxLayout; vs->setSpacing(2);

        auto*vr1=new QHBoxLayout; vr1->setSpacing(2);
        auto*btnSortA =makeSmBtn(Ic::sortAsc(), "Sort A to Z");
        auto*btnSortD =makeSmBtn(Ic::sortDesc(),"Sort Z to A");
        auto*btnFilter=makeSmBtn(Ic::filter(),  "Filter");
        vr1->addWidget(btnSortA); vr1->addWidget(btnSortD); vr1->addWidget(btnFilter);

        auto*vr2=new QHBoxLayout; vr2->setSpacing(2);
        auto*btnFreeze=makeSmBtn(Ic::freeze(),  "Freeze Panes");
        btnFreeze->setMenu(simpleMenu(btnFreeze,{"Freeze Panes","Freeze Top Row","Freeze First Column","---","Unfreeze Panes"}));
        btnFreeze->setPopupMode(QToolButton::MenuButtonPopup);
        auto*btnFind  =makeSmBtn(Ic::findIcon(),"Find & Replace (Ctrl+H)");
        vr2->addWidget(btnFreeze); vr2->addWidget(btnFind); vr2->addStretch();

        vs->addLayout(vr1); vs->addLayout(vr2);

        r1->addWidget(btnSum); r1->addWidget(btnFill); r1->addSpacing(4); r1->addLayout(vs);
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

    tabs->addTab(homeTab,     "Home");
    tabs->addTab(new QWidget, "Insert");
    tabs->addTab(new QWidget, "Page Layout");
    tabs->addTab(new QWidget, "Formulas");
    tabs->addTab(new QWidget, "Data");
    tabs->addTab(new QWidget, "Review");
    tabs->addTab(new QWidget, "View");
    tabs->addTab(new QWidget, "Tools");

    mainLayout->addWidget(tabs);

    auto*bottomLine=new QFrame;
    bottomLine->setFixedHeight(1);
    bottomLine->setStyleSheet("background:#e0e3e8;");
    mainLayout->addWidget(bottomLine);
}

RibbonWidget::~RibbonWidget() { delete d; }

void RibbonWidget::setFormatState(const RibbonFormatState& s) {
    QSignalBlocker b1(d->fontFamily),b2(d->fontSize),
        b3(d->btnBold),b4(d->btnItalic),b5(d->btnUnderline),b6(d->numFormat);
    d->fontFamily->setCurrentFont(QFont(s.fontFamily));
    d->fontSize->setValue(s.fontSize);
    d->btnBold->setChecked(s.bold);
    d->btnItalic->setChecked(s.italic);
    d->btnUnderline->setChecked(s.underline);
    d->btnTextColor->setColor(s.textColor);
    d->btnFillColor->setColor(s.fillColor);
    d->numFormat->setCurrentIndex(s.numberFormat);
    d->btnWrap->setChecked(s.wrapText);
    d->btnAlignL->setChecked(s.hAlign==0);
    d->btnAlignC->setChecked(s.hAlign==1);
    d->btnAlignR->setChecked(s.hAlign==2);
}

extern "C" RIBBON_API RibbonWidget* createRibbonWidget(QWidget* parent) {
    return new RibbonWidget(parent);
}
