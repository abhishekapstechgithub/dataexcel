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

    // ══════════════════ INSERT TAB ═══════════════════════════════════════
    auto* insertTab = new QWidget;
    insertTab->setStyleSheet("background:#ffffff;");
    {
        auto* row = new QHBoxLayout(insertTab);
        row->setContentsMargins(8,2,8,0); row->setSpacing(0);
        auto addS=[&]{ row->addWidget(vSep()); row->addSpacing(4); };

        // Tables group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnPivot = makeLargeBtn(Ic::tableStyle(),"PivotTable","Insert PivotTable",54,70);
            auto* btnPivChart = makeLargeBtn(Ic::tableStyle(),"PivotChart","Insert PivotChart",54,70);
            auto* btnTable = makeLargeBtn(Ic::tableStyle(),"Table","Insert Table (Ctrl+T)",54,70);
            r->addWidget(btnPivot); r->addWidget(btnPivChart); r->addWidget(btnTable);
            col->addLayout(r); col->addWidget(groupLabel("Tables"));
            row->addWidget(grp); addS();
        }
        // Illustrations group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto mkI=[&](const QString& lbl, const QString& tip)->QToolButton*{
                return makeLargeBtn(makeIcon([](QPainter&p,int s){
                    p.setPen(dp(1.5)); p.setBrush(QColor("#e8f5ee"));
                    p.drawRoundedRect(3,3,s-6,s-6,2,2);
                    p.setPen(gp(1.8)); p.drawLine(s/2,5,s/2,s-5);
                    p.drawLine(5,s/2,s-5,s/2);
                }),lbl,tip,54,70);
            };
            r->addWidget(mkI("Pictures","Insert Pictures"));
            auto* btnShapes=makeLargeBtn(makeIcon([](QPainter&p,int s){
                p.setPen(gp(1.8)); p.setBrush(Qt::NoBrush);
                p.drawEllipse(3,3,s-6,s-6);
                p.setPen(dp(1.5)); QPolygon poly; poly<<QPoint(s/2,4)<<QPoint(s-4,s-4)<<QPoint(4,s-4); p.drawPolygon(poly);
            }),"Shapes","Insert Shapes",54,70);
            btnShapes->setMenu(simpleMenu(btnShapes,{"Rectangle","Rounded Rectangle","Circle","Arrow","Line","---","Text Box"}));
            btnShapes->setPopupMode(QToolButton::MenuButtonPopup);
            r->addWidget(btnShapes);
            r->addWidget(mkI("Icons","Insert Icons"));
            r->addWidget(mkI("WordArt","Insert WordArt"));
            col->addLayout(r); col->addWidget(groupLabel("Illustrations"));
            row->addWidget(grp); addS();
        }
        // Charts group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto mkChart=[&](const QString& nm, const QString& sym)->QToolButton*{
                return makeLargeBtn(makeIcon([sym](QPainter&p,int s){
                    QFont f("Segoe UI",int(s*0.5),QFont::Bold); p.setFont(f);
                    p.setPen(QColor("#1a7a45")); p.drawText(QRect(0,0,s,s),Qt::AlignCenter,sym);
                }),nm,"Insert "+nm+" Chart",54,70);
            };
            auto* btnBar     = mkChart("Bar","▬");
            auto* btnLine    = mkChart("Line","⟋");
            auto* btnPie     = mkChart("Pie","◕");
            auto* btnScatter = mkChart("Scatter","⁚");
            r->addWidget(btnBar); r->addWidget(btnLine); r->addWidget(btnPie); r->addWidget(btnScatter);
            col->addLayout(r); col->addWidget(groupLabel("Charts"));
            row->addWidget(grp); addS();
            connect(btnBar,     &QToolButton::clicked,this,[this]{ emit insertChartRequested("Bar"); });
            connect(btnLine,    &QToolButton::clicked,this,[this]{ emit insertChartRequested("Line"); });
            connect(btnPie,     &QToolButton::clicked,this,[this]{ emit insertChartRequested("Pie"); });
            connect(btnScatter, &QToolButton::clicked,this,[this]{ emit insertChartRequested("Scatter"); });
        }
        // Sparklines group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto mkSp=[&](const QString& lbl)->QToolButton*{
                return makeLargeBtn(makeIcon([](QPainter&p,int){
                    p.setPen(gp(1.8));
                    p.drawLine(2,14,6,8); p.drawLine(6,8,10,16); p.drawLine(10,16,14,6); p.drawLine(14,6,18,10);
                }),lbl,"Insert "+lbl+" Sparkline",54,70);
            };
            r->addWidget(mkSp("Line")); r->addWidget(mkSp("Column")); r->addWidget(mkSp("Win/Loss"));
            col->addLayout(r); col->addWidget(groupLabel("Sparklines"));
            row->addWidget(grp); addS();
        }
        // Link + Symbol group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnLink=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(2)); p.setBrush(Qt::NoBrush);
                p.drawArc(2,6,8,8,30*16,300*16); p.drawArc(10,6,8,8,30*16,300*16);
                p.drawLine(7,10,13,10);
            }),"Link","Insert Hyperlink (Ctrl+K)",54,70);
            auto* btnEq=makeLargeBtn(makeIcon([](QPainter&p,int s){
                QFont f("Segoe UI",int(s*0.65)); p.setFont(f); p.setPen(QColor("#1a7a45"));
                p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"√x");
            }),"Equation","Insert Equation",54,70);
            auto* btnSym=makeLargeBtn(makeIcon([](QPainter&p,int s){
                QFont f("Segoe UI",int(s*0.6),QFont::Bold); p.setFont(f); p.setPen(QColor("#444"));
                p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"Ω");
            }),"Symbol","Insert Symbol",54,70);
            r->addWidget(btnLink); r->addWidget(btnEq); r->addWidget(btnSym);
            col->addLayout(r); col->addWidget(groupLabel("Link & Symbol"));
            row->addWidget(grp);
        }
        row->addStretch();
    }

    // ══════════════════ PAGE LAYOUT TAB ══════════════════════════════════
    auto* pageLayoutTab = new QWidget;
    pageLayoutTab->setStyleSheet("background:#ffffff;");
    {
        auto* row = new QHBoxLayout(pageLayoutTab);
        row->setContentsMargins(8,2,8,0); row->setSpacing(0);
        auto addS=[&]{ row->addWidget(vSep()); row->addSpacing(4); };

        // Print Settings group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r1=new QHBoxLayout; r1->setSpacing(2);

            auto* btnPP=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp()); p.setBrush(Qt::white); p.drawRoundedRect(4,2,12,16,1,1);
                p.setPen(gp(1.5)); p.drawLine(6,6,14,6); p.drawLine(6,9,14,9); p.drawLine(6,12,14,12);
                p.setPen(dp(1.8)); p.drawRect(2,12,16,5);
            }),"Print\nPreview","Print Preview",56,70);
            connect(btnPP,&QToolButton::clicked,this,&RibbonWidget::printPreviewRequested);

            auto* btnMarg=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(2)); p.drawLine(4,2,4,18); p.drawLine(16,2,16,18);
                p.setPen(dp(1.2)); p.drawRect(6,4,8,12);
            }),"Margins","Page Margins",44,70);
            btnMarg->setMenu(simpleMenu(btnMarg,{"Normal","Narrow","Wide","---","Custom Margins..."}));
            btnMarg->setPopupMode(QToolButton::MenuButtonPopup);
            connect(btnMarg,&QToolButton::clicked,this,&RibbonWidget::marginsRequested);

            auto* btnOri=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp()); p.setBrush(Qt::white);
                p.drawRoundedRect(5,3,10,14,1,1);
                p.setPen(gp(1.5)); p.drawLine(7,7,13,7); p.drawLine(7,10,13,10);
            }),"Orientation","Page Orientation",44,70);
            btnOri->setMenu(simpleMenu(btnOri,{"Portrait","Landscape"}));
            btnOri->setPopupMode(QToolButton::MenuButtonPopup);
            connect(btnOri->menu(),&QMenu::triggered,this,[this](QAction*a){
                emit orientationChanged(a->text()=="Landscape");
            });

            auto* btnSize=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.drawRoundedRect(4,2,12,16,1,1);
                p.setPen(gp(1.5)); p.drawLine(8,5,12,5); p.drawLine(8,8,12,8);
                p.drawLine(6,11,14,11); p.drawLine(6,14,14,14);
            }),"Size","Paper Size",44,70);
            btnSize->setMenu(simpleMenu(btnSize,{"Letter (8.5×11)","A4 (210×297mm)","A3","Legal","---","More Paper Sizes..."}));
            btnSize->setPopupMode(QToolButton::MenuButtonPopup);

            r1->addWidget(btnPP); r1->addWidget(btnMarg); r1->addWidget(btnOri); r1->addWidget(btnSize);

            auto* r2=new QHBoxLayout; r2->setSpacing(2);
            auto* btnPA=makeMedBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(1.6)); p.drawRect(3,3,14,14);
                p.setPen(dp(1.0)); p.drawLine(8,3,8,17); p.drawLine(3,9,17,9);
            }),"Print Area","Set/Clear Print Area",110,34);
            btnPA->setMenu(simpleMenu(btnPA,{"Set Print Area","Clear Print Area","Add to Print Area"}));
            btnPA->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnPS=makeMedBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.3)); p.drawRect(2,3,16,14);
                p.setPen(gp(1.8)); p.drawLine(10,3,10,17); p.drawLine(2,10,18,10);
            }),"Print Scaling","Fit to Page",110,34);
            btnPS->setMenu(simpleMenu(btnPS,{"Fit Sheet on One Page","Fit All Columns on One Page","Fit All Rows on One Page","---","Custom Scaling..."}));
            btnPS->setPopupMode(QToolButton::MenuButtonPopup);
            r2->addWidget(btnPA); r2->addWidget(btnPS);

            auto* r3=new QHBoxLayout; r3->setSpacing(6);
            auto mkChk=[&](const QString& lbl, const QString& tip)->QToolButton*{
                auto*b=makeSmBtn(makeIcon([](QPainter&p,int){
                    p.setPen(dp(1.3)); p.drawRect(2,5,12,12);
                    p.setPen(gp(2)); p.drawLine(4,11,8,15); p.drawLine(8,15,18,5);
                }),tip,true);
                b->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
                b->setText(lbl); b->setFixedSize(110,22);
                b->setStyleSheet("QToolButton{border:none;background:transparent;font-size:11px;color:#333;}"
                                 "QToolButton:hover{color:#1a7a45;}"
                                 "QToolButton:checked{color:#1a7a45;}");
                return b;
            };
            auto* chkGrid=mkChk("Print Gridlines","Print Gridlines");
            auto* chkHead=mkChk("Print Headings","Print Row/Column Headings");
            chkGrid->setChecked(false); chkHead->setChecked(false);
            r3->addWidget(chkGrid); r3->addWidget(chkHead);

            col->addLayout(r1); col->addLayout(r2); col->addLayout(r3);
            col->addWidget(groupLabel("Print Settings"));
            row->addWidget(grp); addS();
        }

        // Pagination Settings group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnPBP=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(1.5)); p.setBrush(QColor("#e8f5ee")); p.drawRect(3,3,14,14);
                p.setPen(QPen(QColor("#888"),1,Qt::DashLine)); p.drawLine(3,11,17,11);
            }),"Page Break\nPreview","Page Break Preview",56,70);
            auto* btnInsPB=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(2)); p.drawLine(2,10,18,10);
                p.setPen(Qt::NoPen); p.setBrush(QColor("#1a7a45"));
                QPolygon a; a<<QPoint(9,6)<<QPoint(13,6)<<QPoint(11,3); p.drawPolygon(a);
            }),"Insert Page\nBreak","Insert Page Break",56,70);
            auto* btnVPB=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(QPen(QColor("#888"),1,Qt::DashLine)); p.drawLine(10,2,10,18);
                p.setPen(gp(1.5)); p.drawRect(2,4,16,12);
            }),"View Page\nBreak","View Page Break",56,70);
            r->addWidget(btnPBP); r->addWidget(btnInsPB); r->addWidget(btnVPB);
            col->addLayout(r); col->addWidget(groupLabel("Pagination Settings"));
            row->addWidget(grp); addS();
        }

        // Themes group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnTheme=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen); p.drawRect(2,2,8,8);
                p.setBrush(QColor("#2196F3")); p.drawRect(10,2,8,8);
                p.setBrush(QColor("#FF9800")); p.drawRect(2,10,8,8);
                p.setBrush(QColor("#f44336")); p.drawRect(10,10,8,8);
            }),"Themes","Apply Theme",56,70);
            btnTheme->setMenu(simpleMenu(btnTheme,{"Office","Green","Blue","Dark","---","Browse Themes..."}));
            btnTheme->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnBg=makeLargeBtn(makeIcon([](QPainter&p,int){
                QLinearGradient g(0,0,20,20); g.setColorAt(0,QColor("#e8f5ee")); g.setColorAt(1,QColor("#1a7a45"));
                p.fillRect(1,1,18,18,g);
                p.setPen(QPen(Qt::white,1.5)); p.drawLine(5,10,15,10); p.drawLine(10,5,10,15);
            }),"Background","Sheet Background",56,70);
            r->addWidget(btnTheme); r->addWidget(btnBg);
            col->addLayout(r); col->addWidget(groupLabel("Themes"));
            row->addWidget(grp);
        }

        row->addStretch();
    }

    // ══════════════════ FORMULAS TAB ═════════════════════════════════════
    auto* formulasTab = new QWidget;
    formulasTab->setStyleSheet("background:#ffffff;");
    {
        auto* row = new QHBoxLayout(formulasTab);
        row->setContentsMargins(8,2,8,0); row->setSpacing(0);
        auto addS=[&]{ row->addWidget(vSep()); row->addSpacing(4); };

        // Quick Functions group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnInsF=makeLargeBtn(makeIcon([](QPainter&p,int s){
                QFont f("Segoe UI",int(s*0.5)); p.setFont(f); p.setPen(QColor("#1a7a45"));
                p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"fx");
            }),"Insert\nFunction","Insert Function (Shift+F3)",56,70);
            auto* btnSum2=makeLargeBtn(Ic::autosum(),"AutoSum","AutoSum (Alt+=)",56,70);
            btnSum2->setMenu(simpleMenu(btnSum2,{"Sum","Average","Count Numbers","Max","Min","---","More Functions..."}));
            btnSum2->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnRecent=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.5)); p.drawEllipse(2,2,16,16);
                p.drawLine(10,6,10,10); p.drawLine(10,10,14,12);
            }),"Recently\nUsed","Recently Used Functions",56,70);
            btnRecent->setMenu(simpleMenu(btnRecent,{"SUM","IF","VLOOKUP","COUNTIF","AVERAGE","---","More..."}));
            btnRecent->setPopupMode(QToolButton::MenuButtonPopup);
            r->addWidget(btnInsF); r->addWidget(btnSum2); r->addWidget(btnRecent);
            col->addLayout(r); col->addWidget(groupLabel("Quick Functions"));
            row->addWidget(grp); addS();
            connect(btnSum2,&QToolButton::clicked,this,&RibbonWidget::autoSumRequested);
        }

        // Function Library group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto mkFnBtn=[&](const QString& nm, const QString& tip, QStringList fns)->QToolButton*{
                auto*b=makeMedBtn(Ic::autosum(),nm,tip,110,34);
                if(!fns.isEmpty()) b->setMenu(simpleMenu(b,fns));
                return b;
            };
            auto* r1=new QHBoxLayout; r1->setSpacing(3);
            r1->addWidget(mkFnBtn("Financial","Financial Functions",{"PMT","PV","FV","RATE","NPV","IRR","SLN","DB","---","More Financial..."}));
            r1->addWidget(mkFnBtn("Logical","Logical Functions",{"IF","AND","OR","NOT","IFERROR","IFS","SWITCH"}));
            r1->addWidget(mkFnBtn("Text","Text Functions",{"CONCATENATE","LEN","LEFT","RIGHT","MID","FIND","REPLACE","UPPER","LOWER","TRIM","TEXT"}));
            auto* r2=new QHBoxLayout; r2->setSpacing(3);
            r2->addWidget(mkFnBtn("Date & Time","Date/Time Functions",{"TODAY","NOW","DATE","YEAR","MONTH","DAY","DATEDIF","NETWORKDAYS","WEEKDAY"}));
            r2->addWidget(mkFnBtn("Lookup &\nReference","Lookup Functions",{"VLOOKUP","HLOOKUP","INDEX","MATCH","OFFSET","INDIRECT","CHOOSE","XLOOKUP"}));
            r2->addWidget(mkFnBtn("Math\nand Trig","Math Functions",{"SUM","ROUND","ABS","SQRT","POWER","MOD","INT","CEILING","FLOOR","SUMIF","SUMIFS"}));
            auto* r3=new QHBoxLayout; r3->setSpacing(3);
            r3->addWidget(mkFnBtn("More\nFunctions","Statistical & More",{"COUNTIF","COUNTIFS","AVERAGEIF","STDEV","VAR","MEDIAN","PERCENTILE","RANK","---","Statistical...","Engineering...","Information..."}));
            col->addLayout(r1); col->addLayout(r2); col->addLayout(r3);
            col->addWidget(groupLabel("Function Library"));
            row->addWidget(grp); addS();
        }

        // Defined Names group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnNM=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.drawRoundedRect(2,4,16,12,2,2);
                p.setPen(gp(1.5)); p.drawLine(5,8,15,8); p.drawLine(5,12,11,12);
            }),"Name\nManager","Name Manager (Ctrl+F3)",56,70);
            auto* btnCreate=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(2.2)); p.drawLine(10,3,10,17); p.drawLine(3,10,17,10);
            }),"Create","Create from Selection",56,70);
            auto* btnPaste2=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp()); p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(3,6,14,12,1,1);
                p.setPen(gp(1.5)); p.drawLine(6,10,14,10); p.drawLine(10,7,10,13);
            }),"Paste","Paste Name (F3)",56,70);
            r->addWidget(btnNM); r->addWidget(btnCreate); r->addWidget(btnPaste2);
            col->addLayout(r); col->addWidget(groupLabel("Defined Names"));
            row->addWidget(grp); addS();
        }

        // Formula Auditing group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnTrace=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(1.8)); p.drawLine(3,15,10,8); p.drawLine(10,8,17,15);
                p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
                QPolygon a; a<<QPoint(7,8)<<QPoint(13,8)<<QPoint(10,4); p.drawPolygon(a);
                p.setPen(dp(1.2)); p.drawRect(7,14,6,5);
            }),"Trace\nPrecedents","Trace Precedents",56,70);
            auto* btnDep=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(1.8)); p.drawLine(3,5,10,12); p.drawLine(10,12,17,5);
                p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
                QPolygon a; a<<QPoint(7,12)<<QPoint(13,12)<<QPoint(10,16); p.drawPolygon(a);
                p.setPen(dp(1.2)); p.drawRect(7,1,6,5);
            }),"Trace\nDependents","Trace Dependents",56,70);
            auto* btnErr=makeLargeBtn(makeIcon([](QPainter&p,int s){
                QFont f("Segoe UI",int(s*0.5),QFont::Bold); p.setFont(f); p.setPen(QColor("#c0392b"));
                p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"!");
            }),"Error\nChecking","Error Checking",56,70);
            r->addWidget(btnTrace); r->addWidget(btnDep); r->addWidget(btnErr);
            col->addLayout(r); col->addWidget(groupLabel("Formula Auditing"));
            row->addWidget(grp); addS();
        }

        // Calculation group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnCalcNow=makeLargeBtn(makeIcon([](QPainter&p,int s){
                QFont f("Segoe UI",int(s*0.55),QFont::Bold); p.setFont(f); p.setPen(QColor("#1a7a45"));
                p.drawText(QRect(0,0,s,s-4),Qt::AlignCenter,"F9");
                p.setPen(gp(1.5)); p.drawLine(2,s-4,s-2,s-4);
            }),"Calculate\nNow","Calculate All (F9)",56,70);
            auto* btnCalcSht=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.drawRoundedRect(2,3,16,14,1,1);
                p.setPen(gp(1.8)); p.drawLine(5,8,15,8); p.drawLine(5,12,10,12);
                p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
                QPolygon a; a<<QPoint(13,10)<<QPoint(17,12)<<QPoint(13,14); p.drawPolygon(a);
            }),"Calculate\nSheet","Calculate Sheet (Shift+F9)",56,70);
            auto* btnCalcOpt=makeMedBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp()); p.setBrush(Qt::white); p.drawRoundedRect(2,2,16,16,2,2);
                p.setPen(gp(1.5)); p.drawLine(5,7,15,7); p.drawLine(5,10,15,10); p.drawLine(5,13,10,13);
            }),"Calculation\nOptions","Calculation Options",110,40);
            btnCalcOpt->setMenu(simpleMenu(btnCalcOpt,{"Automatic","Automatic Except Tables","Manual","---","Recalculate Before Saving"}));
            btnCalcOpt->setPopupMode(QToolButton::MenuButtonPopup);
            r->addWidget(btnCalcNow); r->addWidget(btnCalcSht);
            col->addLayout(r); col->addWidget(btnCalcOpt);
            col->addWidget(groupLabel("Calculation"));
            row->addWidget(grp);
        }
        row->addStretch();
    }

    // ══════════════════ DATA TAB ══════════════════════════════════════════
    auto* dataTab = new QWidget;
    dataTab->setStyleSheet("background:#ffffff;");
    {
        auto* row = new QHBoxLayout(dataTab);
        row->setContentsMargins(8,2,8,0); row->setSpacing(0);
        auto addS=[&]{ row->addWidget(vSep()); row->addSpacing(4); };

        // PivotTable group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* btnPT=makeLargeBtn(Ic::tableStyle(),"PivotTable","Insert PivotTable",56,70);
            col->addWidget(btnPT); col->addWidget(groupLabel("PivotTable"));
            row->addWidget(grp); addS();
        }

        // Filter & Sort group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r1=new QHBoxLayout; r1->setSpacing(2);
            auto* btnFilt2=makeLargeBtn(Ic::filter(),"Filter","Toggle AutoFilter",54,70);
            auto* btnAll=makeSmBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(2)); p.drawLine(2,5,18,5); p.drawLine(5,9,15,9); p.drawLine(8,13,12,13);
                p.setPen(dp(1.5)); p.drawLine(15,13,18,13);
            }),"Show All - Clear Filters");
            auto* btnReapply=makeSmBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(1.8)); p.drawArc(3,3,14,14,60*16,300*16);
                p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
                QPolygon a; a<<QPoint(15,5)<<QPoint(18,2)<<QPoint(18,8); p.drawPolygon(a);
            }),"Reapply Filter");
            r1->addWidget(btnFilt2); r1->addSpacing(4);
            auto* vs2=new QVBoxLayout; vs2->setSpacing(2);
            vs2->addWidget(btnAll); vs2->addWidget(btnReapply);
            r1->addLayout(vs2);
            auto* r2=new QHBoxLayout; r2->setSpacing(2);
            auto* btnSortA2=makeSmBtn(Ic::sortAsc(), "Sort A to Z");
            auto* btnSortD2=makeSmBtn(Ic::sortDesc(),"Sort Z to A");
            auto* btnSort=makeMedBtn(Ic::sortAsc(),"Sort","Custom Sort",80,28);
            btnSort->setMenu(simpleMenu(btnSort,{"Sort A to Z","Sort Z to A","---","Custom Sort..."}));
            btnSort->setPopupMode(QToolButton::MenuButtonPopup);
            r2->addWidget(btnSortA2); r2->addWidget(btnSortD2); r2->addWidget(btnSort);
            col->addLayout(r1); col->addLayout(r2);
            col->addWidget(groupLabel("Filter & Sort"));
            row->addWidget(grp); addS();
            connect(btnFilt2, &QToolButton::clicked,this,&RibbonWidget::filterRequested);
            connect(btnSortA2,&QToolButton::clicked,this,&RibbonWidget::sortAscRequested);
            connect(btnSortD2,&QToolButton::clicked,this,&RibbonWidget::sortDescRequested);
        }

        // Data Tools group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r1=new QHBoxLayout; r1->setSpacing(2);
            auto mkDT=[&](const QString& nm, const QString& tip)->QToolButton*{
                return makeMedBtn(Ic::fmtCell(),nm,tip,120,34);
            };
            auto* btnHighDup=mkDT("Highlight\nDuplicates","Highlight Duplicate Values");
            btnHighDup->setMenu(simpleMenu(btnHighDup,{"Highlight Duplicates","Highlight Unique Values","---","Clear Highlighting"}));
            btnHighDup->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnMngDup=mkDT("Manage\nDuplicates","Remove Duplicates");
            btnMngDup->setMenu(simpleMenu(btnMngDup,{"Remove Duplicates","Remove Duplicates (Keep Last)","---","Show Duplicates Only"}));
            btnMngDup->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnTxt=mkDT("Text to\nColumns","Text to Columns");
            auto* btnValid=mkDT("Validation","Data Validation");
            btnValid->setMenu(simpleMenu(btnValid,{"Data Validation...","Circle Invalid Data","Clear Validation Circles"}));
            btnValid->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnFill2=mkDT("Fill","Flash Fill (Ctrl+E)");
            btnFill2->setMenu(simpleMenu(btnFill2,{"Flash Fill","Fill Down","Fill Right","Fill Up","Fill Left","---","Series..."}));
            btnFill2->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnConsolidate=mkDT("Consolidate","Consolidate Data");
            auto* btnDropDown=mkDT("Insert Drop-\nDown List","Insert Drop-Down List");
            r1->addWidget(btnHighDup); r1->addWidget(btnMngDup);
            auto* r2=new QHBoxLayout; r2->setSpacing(2);
            r2->addWidget(btnTxt); r2->addWidget(btnValid);
            auto* r3=new QHBoxLayout; r3->setSpacing(2);
            r3->addWidget(btnFill2); r3->addWidget(btnConsolidate);
            auto* r4=new QHBoxLayout; r4->setSpacing(2);
            r4->addWidget(btnDropDown);
            col->addLayout(r1); col->addLayout(r2); col->addLayout(r3); col->addLayout(r4);
            col->addWidget(groupLabel("Data Tools"));
            row->addWidget(grp); addS();
        }

        // Outline group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r1=new QHBoxLayout; r1->setSpacing(2);
            auto* btnGrp=makeMedBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(1.8)); p.drawLine(5,3,15,3); p.drawLine(5,3,5,17); p.drawLine(5,17,15,17);
                p.drawLine(15,3,15,17);
            }),"Group","Group Rows/Columns",110,34);
            btnGrp->setMenu(simpleMenu(btnGrp,{"Group...","Auto Outline"}));
            btnGrp->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnUGrp=makeMedBtn(makeIcon([](QPainter&p,int){
                p.setPen(rp(1.8)); p.drawLine(5,3,15,3); p.drawLine(5,3,5,17); p.drawLine(5,17,15,17);
                p.drawLine(15,3,15,17);
                p.setPen(rp(2)); p.drawLine(7,10,13,10);
            }),"Ungroup","Ungroup Rows/Columns",110,34);
            btnUGrp->setMenu(simpleMenu(btnUGrp,{"Ungroup...","Clear Outline"}));
            btnUGrp->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnSub=makeMedBtn(Ic::autosum(),"Subtotal","Subtotal",110,34);
            r1->addWidget(btnGrp); r1->addWidget(btnUGrp);
            auto* r2=new QHBoxLayout; r2->setSpacing(2);
            auto* btnShowD=makeSmBtn(makeIcon([](QPainter&p,int){ p.setPen(gp(2)); p.drawLine(10,4,10,16); p.drawLine(4,10,16,10); }),"Show Detail");
            auto* btnHideD=makeSmBtn(makeIcon([](QPainter&p,int){ p.setPen(rp(2)); p.drawLine(4,10,16,10); }),"Hide Detail");
            r2->addWidget(btnShowD); r2->addWidget(btnHideD); r2->addWidget(btnSub);
            col->addLayout(r1); col->addLayout(r2);
            col->addWidget(groupLabel("Outline"));
            row->addWidget(grp);
        }

        row->addStretch();
    }

    // ══════════════════ REVIEW TAB ════════════════════════════════════════
    auto* reviewTab = new QWidget;
    reviewTab->setStyleSheet("background:#ffffff;");
    {
        auto* row = new QHBoxLayout(reviewTab);
        row->setContentsMargins(8,2,8,0); row->setSpacing(0);
        auto addS=[&]{ row->addWidget(vSep()); row->addSpacing(4); };

        // Proofing group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* btnSpell=makeLargeBtn(makeIcon([](QPainter&p,int s){
                QFont f("Segoe UI",int(s*0.55),QFont::Bold); p.setFont(f);
                p.setPen(QColor("#444")); p.drawText(QRect(0,0,s,s-4),Qt::AlignCenter,"ABC");
                p.setPen(gp(2)); p.drawLine(2,s-3,s-2,s-3);
            }),"Spelling","Spell Check (F7)",56,70);
            connect(btnSpell,&QToolButton::clicked,this,&RibbonWidget::spellCheckRequested);
            col->addWidget(btnSpell); col->addWidget(groupLabel("Proofing"));
            row->addWidget(grp); addS();
        }

        // Comments group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnNewCmt=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.setBrush(QColor("#fff9e6")); p.drawRoundedRect(2,2,16,12,2,2);
                p.drawLine(6,17,10,14); p.setPen(gp(2)); p.drawLine(5,7,15,7); p.drawLine(5,10,11,10);
            }),"New\nComment","New Comment (Shift+F2)",54,70);
            connect(btnNewCmt,&QToolButton::clicked,this,&RibbonWidget::newCommentRequested);
            auto* vs3=new QVBoxLayout; vs3->setSpacing(2);
            auto* btnDelCmt=makeSmBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.setBrush(QColor("#ffe0e0")); p.drawRoundedRect(2,2,16,12,2,2);
                p.drawLine(6,17,10,14); p.setPen(rp(2)); p.drawLine(6,6,14,12); p.drawLine(14,6,6,12);
            }),"Delete Comment");
            auto* btnPrev=makeSmBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.8)); QPolygon a; a<<QPoint(14,4)<<QPoint(6,10)<<QPoint(14,16); p.drawPolygon(a);
            }),"Previous Comment");
            auto* btnNext=makeSmBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.8)); QPolygon a; a<<QPoint(6,4)<<QPoint(14,10)<<QPoint(6,16); p.drawPolygon(a);
            }),"Next Comment");
            vs3->addWidget(btnDelCmt); vs3->addWidget(btnPrev); vs3->addWidget(btnNext);
            r->addWidget(btnNewCmt); r->addLayout(vs3);
            auto* r2=new QHBoxLayout; r2->setSpacing(2);
            auto* btnShow=makeMedBtn(Ic::fmtCell(),"Show","Show/Hide Comments",80,28);
            btnShow->setMenu(simpleMenu(btnShow,{"Show All Comments","Hide All Comments","---","Show Comment Indicators"}));
            btnShow->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnReset=makeMedBtn(Ic::fmtCell(),"Reset","Reset Comment",80,28);
            r2->addWidget(btnShow); r2->addWidget(btnReset);
            col->addLayout(r); col->addLayout(r2);
            col->addWidget(groupLabel("Comments"));
            row->addWidget(grp); addS();
        }

        // Protect group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto mkLock=[&](const QString& nm, const QString& tip)->QToolButton*{
                return makeLargeBtn(makeIcon([](QPainter&p,int){
                    p.setPen(gp(1.8)); p.setBrush(QColor("#e8f5ee"));
                    p.drawRoundedRect(5,8,10,10,2,2);
                    p.setPen(gp(2)); p.setBrush(Qt::NoBrush);
                    p.drawArc(6,3,8,8,0,180*16);
                }),nm,tip,54,70);
            };
            auto* btnLockCell=mkLock("Lock Cell","Lock/Unlock Cell");
            auto* btnAllowEdit=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.setBrush(QColor("#fff3e0")); p.drawRoundedRect(2,2,16,16,2,2);
                p.setPen(gp(1.5)); p.drawLine(5,8,9,12); p.drawLine(9,12,15,6);
            }),"Allow Edit\nRanges","Allow Users to Edit Ranges",54,70);
            auto* btnProtSht=mkLock("Protect\nSheet","Protect Sheet");
            auto* btnProtWB=mkLock("Protect\nWorkbook","Protect Workbook");
            auto* btnShare=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(1.8)); p.drawEllipse(3,3,6,6); p.drawEllipse(11,3,6,6); p.drawEllipse(7,11,6,6);
                p.drawLine(9,6,12,6); p.drawLine(10,9,13,9);
            }),"Share\nWorkbook","Share Workbook",54,70);
            connect(btnProtSht,&QToolButton::clicked,this,&RibbonWidget::protectSheetRequested);
            connect(btnProtWB, &QToolButton::clicked,this,&RibbonWidget::protectWorkbookRequested);
            r->addWidget(btnLockCell); r->addWidget(btnAllowEdit);
            r->addWidget(btnProtSht); r->addWidget(btnProtWB); r->addWidget(btnShare);
            col->addLayout(r); col->addWidget(groupLabel("Protect"));
            row->addWidget(grp);
        }
        row->addStretch();
    }

    // ══════════════════ VIEW TAB ══════════════════════════════════════════
    auto* viewTab = new QWidget;
    viewTab->setStyleSheet("background:#ffffff;");
    {
        auto* row = new QHBoxLayout(viewTab);
        row->setContentsMargins(8,2,8,0); row->setSpacing(0);
        auto addS=[&]{ row->addWidget(vSep()); row->addSpacing(4); };

        // Workbook Views group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto mkView=[&](const QString& nm, const QString& tip)->QToolButton*{
                return makeLargeBtn(makeIcon([](QPainter&p,int){
                    p.setPen(dp(1.2)); p.drawRect(2,4,16,12);
                    p.setPen(gp(1.5)); p.drawLine(2,8,18,8);
                    p.setPen(dp()); p.drawLine(7,8,7,16);
                }),nm,tip,54,70);
            };
            auto* btnNorm=mkView("Normal","Normal View");
            auto* btnPBP=mkView("Page Break\nPreview","Page Break Preview");
            auto* btnPageL=mkView("Page Layout","Page Layout View");
            r->addWidget(btnNorm); r->addWidget(btnPBP); r->addWidget(btnPageL);
            col->addLayout(r); col->addWidget(groupLabel("Workbook Views"));
            row->addWidget(grp); addS();
        }

        // Show group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto mkChk2=[&](const QString& lbl, bool chk)->QToolButton*{
                auto*b=makeSmBtn(makeIcon([](QPainter&p,int){
                    p.setPen(dp(1.2)); p.drawRect(1,4,10,10);
                    p.setPen(gp(2)); p.drawLine(3,9,6,13); p.drawLine(6,13,11,5);
                }),lbl,true);
                b->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
                b->setText(lbl); b->setChecked(chk); b->setFixedSize(140,24);
                b->setStyleSheet("QToolButton{border:none;background:transparent;font-size:11px;color:#333;text-align:left;}"
                                 "QToolButton:hover{color:#1a7a45;}QToolButton:checked{color:#1a7a45;}");
                return b;
            };
            auto* chkFmtBar  =mkChk2("Formula Bar",true);
            auto* chkGridLines=mkChk2("Gridlines",true);
            auto* chkHeadings=mkChk2("Headings",true);
            auto* chkFullSc  =mkChk2("Full Screen",false);
            auto* chkHiRowCol=mkChk2("Highlight Row & Column",false);
            col->addWidget(chkFmtBar); col->addWidget(chkGridLines); col->addWidget(chkHeadings);
            col->addWidget(chkFullSc); col->addWidget(chkHiRowCol);
            col->addWidget(groupLabel("Show"));
            row->addWidget(grp); addS();
            connect(chkFmtBar,   &QToolButton::toggled,this,&RibbonWidget::showFormulaBarToggled);
            connect(chkGridLines,&QToolButton::toggled,this,&RibbonWidget::showGridlinesToggled);
            connect(chkHeadings, &QToolButton::toggled,this,&RibbonWidget::showHeadingsToggled);
        }

        // Zoom group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* btnZoom=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.8)); p.setBrush(Qt::NoBrush); p.drawEllipse(2,2,12,12);
                p.drawLine(11,11,18,18); p.setPen(gp(1.8)); p.drawLine(6,8,10,8); p.drawLine(8,6,8,10);
            }),"Zoom","Open Zoom Dialog",56,70);
            connect(btnZoom,&QToolButton::clicked,this,[this]{ emit zoomToValueRequested(100); });
            auto* r2=new QHBoxLayout; r2->setSpacing(2);
            auto mkZ=[&](int pct)->QToolButton*{
                auto*b=makeSmBtn(makeIcon([pct](QPainter&p,int s){
                    QFont f("Segoe UI",int(s*0.38)); p.setFont(f); p.setPen(QColor("#1a7a45"));
                    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,QString::number(pct)+"%");
                }),QString("Set Zoom to %1%").arg(pct));
                connect(b,&QToolButton::clicked,[this,pct]{ emit zoomToValueRequested(pct); });
                return b;
            };
            r2->addWidget(mkZ(75)); r2->addWidget(mkZ(100)); r2->addWidget(mkZ(150));
            col->addWidget(btnZoom); col->addLayout(r2);
            col->addWidget(groupLabel("Zoom"));
            row->addWidget(grp); addS();
        }

        // Window group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnFrz=makeLargeBtn(Ic::freeze(),"Freeze\nPanes","Freeze Panes",54,70);
            btnFrz->setMenu(simpleMenu(btnFrz,{"Freeze Panes","Freeze Top Row","Freeze First Column","---","Unfreeze Panes"}));
            btnFrz->setPopupMode(QToolButton::MenuButtonPopup);
            connect(btnFrz->menu(),&QMenu::triggered,this,[this](QAction*a){
                QString t=a->text();
                if(t=="Freeze Panes")        emit freezePanesRequested("panes");
                else if(t=="Freeze Top Row") emit freezePanesRequested("row");
                else if(t=="Freeze First Column") emit freezePanesRequested("col");
                else if(t=="Unfreeze Panes") emit freezePanesRequested("unfreeze");
            });
            auto* btnSplit=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(gp(2)); p.drawLine(10,2,10,18); p.drawLine(2,10,18,10);
                p.setPen(dp(1.0)); p.drawRect(2,2,8,8); p.drawRect(10,2,8,8);
                p.drawRect(2,10,8,8); p.drawRect(10,10,8,8);
            }),"Split\nWindow","Split Window",54,70);
            auto* btnArrange=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.drawRect(1,1,10,10); p.drawRect(9,9,10,10);
            }),"Arrange\nAll","Arrange Windows",54,70);
            btnArrange->setMenu(simpleMenu(btnArrange,{"Tiled","Horizontal","Vertical","Cascade"}));
            btnArrange->setPopupMode(QToolButton::MenuButtonPopup);
            auto* btnNewWin=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.drawRect(1,3,14,13);
                p.drawRect(5,1,14,13);
                p.setPen(gp(2)); p.drawLine(12,6,12,10); p.drawLine(10,8,14,8);
            }),"New\nWindow","New Window",54,70);
            r->addWidget(btnFrz); r->addWidget(btnSplit); r->addWidget(btnArrange); r->addWidget(btnNewWin);
            col->addLayout(r); col->addWidget(groupLabel("Window"));
            row->addWidget(grp);
        }
        row->addStretch();
    }

    // ══════════════════ TOOLS TAB ═════════════════════════════════════════
    auto* toolsTab = new QWidget;
    toolsTab->setStyleSheet("background:#ffffff;");
    {
        auto* row = new QHBoxLayout(toolsTab);
        row->setContentsMargins(8,2,8,0); row->setSpacing(0);
        auto addS=[&]{ row->addWidget(vSep()); row->addSpacing(4); };

        // Developer group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto mkDev=[&](const QString& nm, const QString& tip)->QToolButton*{
                return makeLargeBtn(makeIcon([](QPainter&p,int){
                    p.setPen(gp(1.8)); p.drawLine(3,10,8,5); p.drawLine(3,10,8,15);
                    p.drawLine(17,10,12,5); p.drawLine(17,10,12,15);
                }),nm,tip,54,70);
            };
            r->addWidget(mkDev("VB Macros","Run VB Macros"));
            r->addWidget(mkDev("Security","Macro Security Settings"));
            r->addWidget(mkDev("VB Editor","Open VB Editor"));
            auto* r2=new QHBoxLayout; r2->setSpacing(2);
            auto mkDev2=[&](const QString& nm, const QString& tip)->QToolButton*{
                return makeMedBtn(makeIcon([](QPainter&p,int){
                    p.setPen(dp(1.3)); p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(2,4,16,12,2,2);
                    p.setPen(gp(1.5)); p.drawLine(5,8,7,10); p.drawLine(7,10,5,12);
                    p.setPen(dp(1.2)); p.drawLine(9,7,15,7); p.drawLine(9,10,15,10); p.drawLine(9,13,13,13);
                }),nm,tip,110,34);
            };
            r2->addWidget(mkDev2("Python Script\nEditor","Open Python Script Editor"));
            auto* r3=new QHBoxLayout; r3->setSpacing(2);
            r3->addWidget(mkDev2("Add-ins","Manage Add-ins"));
            r3->addWidget(mkDev2("COM Add-ins","COM Add-ins"));
            col->addLayout(r); col->addLayout(r2); col->addLayout(r3);
            col->addWidget(groupLabel("Developer"));
            row->addWidget(grp); addS();
        }

        // Input group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto* btnForm=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp(1.2)); p.setBrush(Qt::white); p.drawRoundedRect(2,2,16,16,2,2);
                p.setPen(gp(1.5)); p.drawLine(5,7,10,7); p.drawLine(5,11,8,11);
                p.setPen(dp(1.2)); p.drawRoundedRect(10,9,6,5,1,1);
            }),"Form","Insert Form Controls",56,70);
            auto* btnInvoice=makeLargeBtn(makeIcon([](QPainter&p,int){
                p.setPen(dp()); p.setBrush(QColor("#fff9e6")); p.drawRoundedRect(2,2,16,16,1,1);
                p.setPen(gp(1.5)); p.drawLine(5,6,15,6); p.drawLine(5,9,15,9); p.drawLine(5,12,10,12);
                p.setPen(QColor("#f0c040")); p.drawLine(5,15,15,15);
            }),"Invoice\nMaker","Create Invoice",56,70);
            r->addWidget(btnForm); r->addWidget(btnInvoice);
            col->addLayout(r); col->addWidget(groupLabel("Input"));
            row->addWidget(grp); addS();
        }

        // Data Processing group
        {
            auto* grp=new QWidget; auto* col=new QVBoxLayout(grp);
            col->setContentsMargins(2,0,2,0); col->setSpacing(2);
            auto* r=new QHBoxLayout; r->setSpacing(2);
            auto mkDP=[&](const QString& nm, const QString& tip, QStringList items)->QToolButton*{
                auto*b=makeLargeBtn(makeIcon([](QPainter&p,int){
                    p.setPen(dp(1.2)); p.drawRect(2,2,7,16); p.drawRect(11,2,7,16);
                    p.setPen(gp(2)); p.drawLine(9,10,11,10);
                    p.setBrush(QColor("#1a7a45")); p.setPen(Qt::NoPen);
                    QPolygon a; a<<QPoint(9,8)<<QPoint(9,12)<<QPoint(12,10); p.drawPolygon(a);
                }),nm,tip,54,70);
                if(!items.isEmpty()){ b->setMenu(simpleMenu(b,items)); b->setPopupMode(QToolButton::MenuButtonPopup); }
                return b;
            };
            auto* btnSplitSh=mkDP("Split\nSheets","Split Sheet by Column",{"Split by Column Value","Split by Row Count","---","Split to Workbooks"});
            auto* btnMergeSh=mkDP("Merge\nSheets","Merge Sheets",{"Merge Sheets into One","Merge Workbooks","---","Merge by Column Key"});
            auto* btnEnhTbl=mkDP("Enhance\nTable","Enhance Table",{"Auto-format Table","Add Total Row","Add Filter Row","---","Convert to Data Range"});
            auto* btnQF=makeLargeBtn(makeIcon([](QPainter&p,int s){
                QFont f("Segoe UI",int(s*0.5),QFont::Bold); p.setFont(f); p.setPen(QColor("#1a7a45"));
                p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"QF");
            }),"Quick\nFormula","Quick Formula Templates",54,70);
            btnQF->setMenu(simpleMenu(btnQF,{"Sum Column","Average Column","Count Non-Empty","Rank Values","Running Total","---","More Templates..."}));
            btnQF->setPopupMode(QToolButton::MenuButtonPopup);
            r->addWidget(btnSplitSh); r->addWidget(btnMergeSh); r->addWidget(btnEnhTbl); r->addWidget(btnQF);
            col->addLayout(r); col->addWidget(groupLabel("Data Processing"));
            row->addWidget(grp);
        }
        row->addStretch();
    }

    tabs->addTab(homeTab,       "Home");
    tabs->addTab(insertTab,     "Insert");
    tabs->addTab(pageLayoutTab, "Page Layout");
    tabs->addTab(formulasTab,   "Formulas");
    tabs->addTab(dataTab,       "Data");
    tabs->addTab(reviewTab,     "Review");
    tabs->addTab(viewTab,       "View");
    tabs->addTab(toolsTab,      "Tools");

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
