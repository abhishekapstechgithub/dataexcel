// ═══════════════════════════════════════════════════════════════════════════════
//  MainWindow.cpp — OpenSheet Application Shell
//  Architecture inspired by WPS Office (et.exe + etcore.dll + ksoui.dll pattern)
//  Layers:
//    Shell (MainWindow) → RibbonUI → FormulaBar → GridView → SheetBar → StatusBar
// ═══════════════════════════════════════════════════════════════════════════════
#include "MainWindow.h"
#include "SpreadsheetView.h"
#include "FindReplaceDialog.h"
#include "FormatCellsDialog.h"
#include "ISpreadsheetCore.h"
#include "IFileLoader.h"
#include "RibbonWidget.h"
#include "SpreadsheetTableModel.h"
#include <QMenuBar>
#include <QStatusBar>
#include <QTabBar>
#include <QLabel>
#include <QLineEdit>
#include <QProgressBar>
#include <QFileDialog>
#include <QMessageBox>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QInputDialog>
#include <QKeySequence>
#include <QtConcurrent>
#include <QCloseEvent>
#include <QResizeEvent>
#include <QApplication>
#include <QToolButton>
#include <QSlider>
#include <QFrame>
#include <QPainter>
#include <QPixmap>
#include <QStyle>
#include <QScreen>

// ── Small QPainter icon factory (same as ribbon) ─────────────────────────────
static QIcon shellIcon(std::function<void(QPainter&,int)> fn, int sz=16) {
    QPixmap pm(sz,sz); pm.fill(Qt::transparent);
    QPainter p(&pm); p.setRenderHint(QPainter::Antialiasing); fn(p,sz);
    return QIcon(pm);
}

// ── Constructor ───────────────────────────────────────────────────────────────
MainWindow::MainWindow(QWidget* parent) : QMainWindow(parent) {
    setWindowTitle("OpenSheet");

    // Size to 80% of screen
    QRect screen = QApplication::primaryScreen()->availableGeometry();
    resize(qMin(1366, (int)(screen.width() * 0.85)),
           qMin(768,  (int)(screen.height() * 0.85)));
    setMinimumSize(900, 600);

    // Green title bar menu (like WPS)
    menuBar()->setStyleSheet(
        "QMenuBar { background: #1e7145; color: white; font-size: 12px; "
        "           font-family: 'Segoe UI', Arial; padding: 0 6px; }"
        "QMenuBar::item { padding: 5px 12px; border-radius: 3px; }"
        "QMenuBar::item:selected { background: #155835; }"
        "QMenuBar::item:pressed  { background: #0f4228; }"
        "QMenu { background:#fff; color:#222; border:1px solid #ddd; "
        "        font-size:12px; font-family:'Segoe UI',Arial; }"
        "QMenu::item { padding:5px 30px 5px 14px; }"
        "QMenu::item:selected { background:#e8f5ee; color:#1a6b35; }"
        "QMenu::separator { height:1px; background:#ececec; margin:3px 0; }"
    );

    m_core   = createSpreadsheetCore();
    m_loader = createFileLoader();

    // ── Central widget ─────────────────────────────────────────────────────
    auto* central = new QWidget;
    central->setStyleSheet("background:#ffffff;");
    auto* vl = new QVBoxLayout(central);
    vl->setContentsMargins(0,0,0,0);
    vl->setSpacing(0);

    // 1. Quick Access Toolbar row (inside menu bar)
    buildMenuBar();

    // 2. Ribbon
    m_ribbon = createRibbonWidget(this);
    vl->addWidget(m_ribbon);

    // 2b. Notification bar (dismissible promo/info bar like WPS)
    m_notifBar = buildNotifBar();
    vl->addWidget(m_notifBar);

    // 3. Formula bar
    {
        auto* fb = buildFormulaBar();
        vl->addWidget(fb);
    }

    // 4. Grid
    m_view = new SpreadsheetView(m_core, m_core->sheets().first(), this);
    m_view->setObjectName("gridView");
    vl->addWidget(m_view, 1);

    // 5. Sheet tab bar
    {
        auto* sb = buildSheetBar();
        vl->addWidget(sb);
    }

    setCentralWidget(central);
    buildStatusBar();
    connectRibbon();
    connectView();
    rebuildSheetTabs();
    updateTitle();
    m_view->setFocus();
}

MainWindow::~MainWindow() { delete m_core; delete m_loader; }

// ═══════════════════════════════════════════════════════════════════════════════
//  MENU BAR
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::buildMenuBar() {
    auto* mb = menuBar();

    // File
    auto* fm = mb->addMenu("File");
    fm->addAction("New",      this, &MainWindow::newFile,    QKeySequence::New);
    fm->addAction("Open...",  this, &MainWindow::openFile,   QKeySequence::Open);
    fm->addSeparator();
    fm->addAction("Save",     this, &MainWindow::saveFile,   QKeySequence::Save);
    fm->addAction("Save As...",this,&MainWindow::saveFileAs, QKeySequence::SaveAs);
    fm->addSeparator();
    fm->addAction("Exit",     qApp, &QApplication::quit,     QKeySequence::Quit);

    // Edit
    auto* em = mb->addMenu("Edit");
    em->addAction("Undo", m_core, &ISpreadsheetCore::undo, QKeySequence::Undo);
    em->addAction("Redo", m_core, &ISpreadsheetCore::redo, QKeySequence::Redo);
    em->addSeparator();
    em->addAction("Find && Replace...", this, [this]{
        auto* d = new FindReplaceDialog(m_core, currentSheet(), this);
        d->setAttribute(Qt::WA_DeleteOnClose); d->show();
    }, QKeySequence(Qt::CTRL|Qt::Key_H));

    // Sheet
    auto* sm = mb->addMenu("Sheet");
    sm->addAction("Insert Sheet", this, &MainWindow::addSheet);
    sm->addAction("Delete Sheet", this, &MainWindow::removeSheet);
    sm->addAction("Rename Sheet", this, [this]{ renameSheet(m_sheetBar->currentIndex()); });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  NOTIFICATION BAR  (WPS-style dismissible info/promo bar)
// ═══════════════════════════════════════════════════════════════════════════════
QWidget* MainWindow::buildNotifBar() {
    auto* bar = new QWidget;
    bar->setObjectName("notifBar");
    bar->setFixedHeight(28);
    bar->setStyleSheet(
        "QWidget#notifBar {"
        "  background: qlineargradient(x1:0,y1:0,x2:1,y2:0,"
        "    stop:0 #e8f5ee, stop:1 #f0fff5);"
        "  border-bottom: 1px solid #b8d9c4;"
        "}"
    );

    auto* hl = new QHBoxLayout(bar);
    hl->setContentsMargins(12,3,8,3);
    hl->setSpacing(8);

    // Star icon
    auto* star = new QLabel("★");
    star->setStyleSheet("color:#f0a500; font-size:14px;");
    hl->addWidget(star);

    // Message
    auto* msg = new QLabel(
        "<span style='color:#1a6b35; font-size:11px; font-family:Segoe UI,Arial;'>"
        "<b>OpenSheet Pro</b> — Unlock advanced features: pivot tables, macros, real-time collaboration</span>"
    );
    msg->setTextFormat(Qt::RichText);
    hl->addWidget(msg, 1);

    // Upgrade button
    auto* upgradeBtn = new QToolButton;
    upgradeBtn->setText("Upgrade Now");
    upgradeBtn->setStyleSheet(
        "QToolButton {"
        "  background: #1e7145; color: white; border: none; border-radius: 4px;"
        "  padding: 3px 14px; font-size: 11px; font-weight: 600;"
        "  font-family: 'Segoe UI', Arial;"
        "}"
        "QToolButton:hover { background: #155835; }"
        "QToolButton:pressed { background: #0f4228; }"
    );
    hl->addWidget(upgradeBtn);

    // Dismiss X button
    auto* closeBtn = new QToolButton;
    closeBtn->setText("✕");
    closeBtn->setFixedSize(20,20);
    closeBtn->setStyleSheet(
        "QToolButton { border:none; background:transparent; color:#888; font-size:12px; }"
        "QToolButton:hover { color:#c0392b; background:#ffe8e8; border-radius:3px; }"
    );
    hl->addWidget(closeBtn);

    connect(closeBtn, &QToolButton::clicked, bar, [bar]{ bar->hide(); });
    connect(upgradeBtn, &QToolButton::clicked, this, [this]{
        QMessageBox::information(this, "OpenSheet Pro",
            "Thank you for your interest!\n\nOpenSheet Pro features:\n"
            "• Pivot Tables & Charts\n• VBA Macros\n• Real-time Collaboration\n"
            "• 500GB file support\n• Cloud Storage Integration");
    });

    return bar;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  FORMULA BAR  (WPS-style: name box | ▼ | fx icon | input)
// ═══════════════════════════════════════════════════════════════════════════════
QWidget* MainWindow::buildFormulaBar() {
    auto* bar = new QWidget;
    bar->setObjectName("formulaBarWidget");
    bar->setFixedHeight(30);
    bar->setStyleSheet(
        "QWidget#formulaBarWidget { background:#f5f5f5; border-bottom: 1px solid #d8d8d8; }"
    );

    auto* hl = new QHBoxLayout(bar);
    hl->setContentsMargins(4, 3, 4, 3);
    hl->setSpacing(0);

    // ── Name box ─────────────────────────────────────────────────────────
    m_nameBox = new QLineEdit("A1");
    m_nameBox->setFixedWidth(70);
    m_nameBox->setFixedHeight(22);
    m_nameBox->setAlignment(Qt::AlignCenter);
    m_nameBox->setStyleSheet(
        "QLineEdit { border:1px solid #c8c8c8; border-radius:2px; background:white; "
        "font-size:12px; font-family:'Segoe UI',Arial; padding:0 3px; }"
        "QLineEdit:focus { border-color:#1e7145; }"
    );
    hl->addWidget(m_nameBox);

    // Small dropdown arrow button next to name box (WPS has this)
    auto* nameDropBtn = new QToolButton;
    nameDropBtn->setFixedSize(16, 22);
    nameDropBtn->setStyleSheet(
        "QToolButton { border:1px solid #c8c8c8; border-left:none; border-radius:0 2px 2px 0; "
        "background:#f0f0f0; }"
        "QToolButton:hover { background:#e0e0e0; }"
    );
    nameDropBtn->setArrowType(Qt::DownArrow);
    hl->addWidget(nameDropBtn);
    hl->addSpacing(4);

    // ── Thin separator ────────────────────────────────────────────────────
    auto* vsep = new QFrame;
    vsep->setFrameShape(QFrame::VLine);
    vsep->setFixedHeight(18);
    vsep->setStyleSheet("color:#c8c8c8;");
    hl->addWidget(vsep);
    hl->addSpacing(4);

    // ── fx button ─────────────────────────────────────────────────────────
    m_fxLabel = new QLabel;
    m_fxLabel->setFixedSize(32, 22);
    m_fxLabel->setAlignment(Qt::AlignCenter);
    // Draw "fx" as a proper formula icon
    {
        QPixmap pm(32, 22); pm.fill(Qt::transparent);
        QPainter p(&pm); p.setRenderHint(QPainter::Antialiasing);
        QFont f("Times New Roman", 11, QFont::Bold, true);
        p.setFont(f); p.setPen(QColor("#1e7145"));
        p.drawText(QRect(0,0,32,22), Qt::AlignCenter, "fx");
        m_fxLabel->setPixmap(pm);
    }
    m_fxLabel->setStyleSheet(
        "QLabel { border:1px solid #c8c8c8; border-radius:2px; background:white; }"
        "QLabel:hover { background:#e8f5ee; border-color:#1e7145; }"
    );
    m_fxLabel->setCursor(Qt::PointingHandCursor);
    m_fxLabel->setToolTip("Insert Function");
    hl->addWidget(m_fxLabel);
    hl->addSpacing(4);

    // ── Another separator ─────────────────────────────────────────────────
    auto* vsep2 = new QFrame;
    vsep2->setFrameShape(QFrame::VLine);
    vsep2->setFixedHeight(18);
    vsep2->setStyleSheet("color:#c8c8c8;");
    hl->addWidget(vsep2);
    hl->addSpacing(3);

    // ── Formula / value input ─────────────────────────────────────────────
    m_formulaBar = new QLineEdit;
    m_formulaBar->setFixedHeight(22);
    m_formulaBar->setStyleSheet(
        "QLineEdit { border:1px solid #c8c8c8; border-radius:2px; background:white; "
        "font-size:12px; font-family:'Segoe UI',Arial; padding:0 6px; }"
        "QLineEdit:focus { border-color:#1e7145; }"
    );
    hl->addWidget(m_formulaBar, 1);

    // Connections
    connect(m_formulaBar, &QLineEdit::returnPressed, this, &MainWindow::onFormulaBarReturn);
    connect(m_formulaBar, &QLineEdit::textEdited,    this, [this](const QString&){
        m_formulaEditing = true;
        if (m_statusReady) m_statusReady->setText("Edit");
    });
    connect(m_nameBox, &QLineEdit::returnPressed, this, [this]{
        QString ref = m_nameBox->text().trimmed().toUpper();
        QRegularExpression re("^([A-Z]+)(\\d+)$");
        auto m = re.match(ref);
        if (m.hasMatch()) {
            int col = SpreadsheetTableModel::columnIndex(m.captured(1));
            int row = m.captured(2).toInt() - 1;
            m_view->setCurrentIndex(m_view->model()->index(row, col));
            m_view->setFocus();
        }
    });

    return bar;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SHEET TAB BAR  (WPS-style bottom bar)
// ═══════════════════════════════════════════════════════════════════════════════
QWidget* MainWindow::buildSheetBar() {
    auto* bar = new QWidget;
    bar->setObjectName("sheetBarWidget");
    bar->setFixedHeight(30);
    bar->setStyleSheet(
        "QWidget#sheetBarWidget { background:#f0f0f0; border-top:1px solid #d0d0d0; }"
    );

    auto* hl = new QHBoxLayout(bar);
    hl->setContentsMargins(4,2,4,2);
    hl->setSpacing(2);

    // Nav arrows (WPS style: ◀◀ ◀ ▶ ▶▶)
    auto mkNav = [&](const QString& ch, const QString& tip) {
        auto* b = new QToolButton;
        b->setText(ch); b->setToolTip(tip);
        b->setFixedSize(20,20);
        b->setStyleSheet(
            "QToolButton { border:1px solid #c8c8c8; border-radius:2px; background:#e8e8e8; "
            "font-size:9px; color:#555; }"
            "QToolButton:hover { background:#d0d0d0; color:#1a6b35; }"
            "QToolButton:pressed { background:#c0c0c0; }"
        );
        return b;
    };
    auto* btnFirst = mkNav("◀◀","First sheet");
    auto* btnPrev  = mkNav("◀", "Previous sheet");
    auto* btnNext  = mkNav("▶", "Next sheet");
    auto* btnLast  = mkNav("▶▶","Last sheet");

    hl->addWidget(btnFirst); hl->addWidget(btnPrev);
    hl->addWidget(btnNext);  hl->addWidget(btnLast);
    hl->addSpacing(6);

    // Sheet tabs
    m_sheetBar = new QTabBar;
    m_sheetBar->setExpanding(false);
    m_sheetBar->setMovable(true);
    m_sheetBar->setDrawBase(false);
    m_sheetBar->setElideMode(Qt::ElideRight);
    m_sheetBar->setFixedHeight(26);
    m_sheetBar->setStyleSheet(
        "QTabBar { background:transparent; }"
        "QTabBar::tab {"
        "  min-width:80px; max-width:180px; height:24px;"
        "  padding:0 14px; margin-right:2px; font-size:11px;"
        "  font-family:'Segoe UI',Arial;"
        "  border:1px solid #c0c0c0; border-bottom:none;"
        "  border-radius:3px 3px 0 0;"
        "  background:#e4e4e4; color:#555;"
        "}"
        "QTabBar::tab:selected {"
        "  background:#ffffff; color:#1a6b35; font-weight:600;"
        "  border-bottom:2px solid #1e7145;"
        "}"
        "QTabBar::tab:hover:!selected { background:#efefef; color:#333; }"
    );
    hl->addWidget(m_sheetBar, 1);

    // Add sheet "+" button
    m_addSheetBtn = new QToolButton;
    m_addSheetBtn->setText("+");
    m_addSheetBtn->setToolTip("Insert Sheet");
    m_addSheetBtn->setFixedSize(24,24);
    m_addSheetBtn->setStyleSheet(
        "QToolButton { border:1px solid #c8c8c8; border-radius:12px; "
        "background:#e8e8e8; font-size:14px; color:#666; font-weight:bold; }"
        "QToolButton:hover { background:#1e7145; color:white; border-color:#1e7145; }"
    );
    hl->addWidget(m_addSheetBtn);

    // Connections
    connect(m_sheetBar,   &QTabBar::currentChanged,      this, &MainWindow::switchSheet);
    connect(m_sheetBar,   &QTabBar::tabBarDoubleClicked,  this, &MainWindow::renameSheet);
    connect(m_addSheetBtn,&QToolButton::clicked,          this, &MainWindow::addSheet);
    connect(btnFirst,&QToolButton::clicked,this,[this]{ m_sheetBar->setCurrentIndex(0); });
    connect(btnLast, &QToolButton::clicked,this,[this]{ m_sheetBar->setCurrentIndex(m_sheetBar->count()-1); });
    connect(btnPrev, &QToolButton::clicked,this,[this]{ m_sheetBar->setCurrentIndex(qMax(0,m_sheetBar->currentIndex()-1)); });
    connect(btnNext, &QToolButton::clicked,this,[this]{ m_sheetBar->setCurrentIndex(qMin(m_sheetBar->count()-1,m_sheetBar->currentIndex()+1)); });

    // Right-click context menu
    m_sheetBar->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_sheetBar,&QWidget::customContextMenuRequested,this,[this](const QPoint& pos){
        int idx = m_sheetBar->tabAt(pos);
        QMenu menu(this);
        menu.setStyleSheet("QMenu{font-size:12px;} QMenu::item{padding:5px 20px;} QMenu::item:selected{background:#e8f5ee;color:#1a6b35;}");
        menu.addAction("Insert Sheet",  this, &MainWindow::addSheet);
        if (idx >= 0) {
            menu.addAction("Delete Sheet",  this, &MainWindow::removeSheet);
            menu.addAction("Rename Sheet",  [this,idx]{ renameSheet(idx); });
            menu.addSeparator();
            menu.addAction("Move/Copy Sheet...", this, [this]{ QMessageBox::information(this,"Move/Copy","Not yet implemented"); });
        }
        menu.exec(m_sheetBar->mapToGlobal(pos));
    });

    return bar;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  STATUS BAR  (WPS-style: Ready | mode | stats ... zoom)
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::buildStatusBar() {
    auto* sb = statusBar();
    sb->setSizeGripEnabled(false);
    sb->setStyleSheet(
        "QStatusBar { background:#f0f0f0; border-top:1px solid #d0d0d0; "
        "font-size:11px; font-family:'Segoe UI',Arial; color:#555; min-height:24px; }"
        "QStatusBar::item { border:none; }"
    );

    // Left: small green dot + "Ready"
    auto* readyWidget = new QWidget;
    readyWidget->setStyleSheet("background:transparent;");
    auto* readyHL = new QHBoxLayout(readyWidget);
    readyHL->setContentsMargins(6,0,6,0); readyHL->setSpacing(5);

    auto* greenDot = new QLabel;
    {
        QPixmap pm(10,10); pm.fill(Qt::transparent);
        QPainter p(&pm); p.setRenderHint(QPainter::Antialiasing);
        p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen);
        p.drawEllipse(1,1,8,8);
        greenDot->setPixmap(pm);
    }
    m_statusReady = new QLabel("Ready");
    m_statusReady->setStyleSheet("color:#444; font-size:11px;");
    readyHL->addWidget(greenDot);
    readyHL->addWidget(m_statusReady);
    sb->addWidget(readyWidget);

    // Thin separator
    auto* sep1 = new QFrame; sep1->setFrameShape(QFrame::VLine);
    sep1->setFixedHeight(14); sep1->setStyleSheet("color:#c8c8c8;");
    sb->addWidget(sep1);

    // Cell mode
    m_statusMode = new QLabel("Normal");
    m_statusMode->setStyleSheet("color:#666; font-size:11px; padding:0 8px;");
    sb->addWidget(m_statusMode);

    // Progress bar (hidden)
    m_progress = new QProgressBar;
    m_progress->setFixedSize(160,12);
    m_progress->setVisible(false);
    m_progress->setTextVisible(false);
    m_progress->setStyleSheet(
        "QProgressBar { border:1px solid #c0c0c0; border-radius:2px; background:#e8e8e8; }"
        "QProgressBar::chunk { background:#1e7145; border-radius:1px; }"
    );
    sb->addWidget(m_progress);

    // ── RIGHT SIDE ──────────────────────────────────────────────────────────

    // Selection stats (Average / Count / Sum) — WPS shows these prominently
    m_statusStats = new QLabel;
    m_statusStats->setStyleSheet("color:#444; font-size:11px; padding:0 12px;");
    m_statusStats->setAlignment(Qt::AlignRight|Qt::AlignVCenter);
    sb->addPermanentWidget(m_statusStats);

    // Separator
    auto* sep2 = new QFrame; sep2->setFrameShape(QFrame::VLine);
    sep2->setFixedHeight(14); sep2->setStyleSheet("color:#c8c8c8;");
    sb->addPermanentWidget(sep2);

    // View mode buttons (Normal / Page Layout / Page Break preview) — WPS-style icons
    auto mkViewBtn = [&](const QIcon& icon, const QString& tip) {
        auto* btn = new QToolButton;
        btn->setIcon(icon); btn->setIconSize(QSize(16,16));
        btn->setToolTip(tip); btn->setCheckable(true);
        btn->setFixedSize(22,22); btn->setAutoRaise(true);
        btn->setStyleSheet(
            "QToolButton { border:1px solid transparent; border-radius:2px; background:transparent; }"
            "QToolButton:hover { border-color:#c0c0c0; background:#e0e0e0; }"
            "QToolButton:checked { background:#c8e6d4; border-color:#1e7145; }"
        );
        return btn;
    };

    auto* btnNormal = mkViewBtn(shellIcon([](QPainter& p,int s){
        p.setPen(QPen(QColor("#555"),1)); p.setBrush(Qt::NoBrush);
        p.drawRect(2,2,s-4,s-4);
        p.setPen(QPen(QColor("#555"),0.8));
        p.drawLine(2,s/3+2,s-2,s/3+2); p.drawLine(2,2*s/3+2,s-2,2*s/3+2);
        p.drawLine(s/3+2,2,s/3+2,s-2); p.drawLine(2*s/3+2,2,2*s/3+2,s-2);
    }), "Normal View");

    auto* btnPage = mkViewBtn(shellIcon([](QPainter& p,int s){
        p.setPen(QPen(QColor("#555"),1)); p.setBrush(QColor("#e8e8e8"));
        p.drawRect(3,2,s-8,s-4);
        p.setBrush(Qt::white); p.drawRect(2,1,s-8,s-4);
        p.setPen(QPen(QColor("#555"),0.7));
        p.drawLine(2,5,s-6,5); p.drawLine(2,8,s-6,8); p.drawLine(2,11,s-6,11);
    }), "Page Layout View");

    auto* btnBreak = mkViewBtn(shellIcon([](QPainter& p,int s){
        p.setPen(QPen(QColor("#555"),1)); p.setBrush(Qt::NoBrush);
        p.drawRect(2,2,s-4,s-4);
        p.setPen(QPen(QColor("#1e7145"),1.5,Qt::DashLine));
        p.drawLine(s/2,2,s/2,s-2); p.drawLine(2,s/2,s-2,s/2);
    }), "Page Break Preview");

    btnNormal->setChecked(true);
    sb->addPermanentWidget(btnNormal);
    sb->addPermanentWidget(btnPage);
    sb->addPermanentWidget(btnBreak);

    // Separator
    auto* sep3 = new QFrame; sep3->setFrameShape(QFrame::VLine);
    sep3->setFixedHeight(14); sep3->setStyleSheet("color:#c8c8c8;");
    sb->addPermanentWidget(sep3);

    // Zoom percentage label
    m_zoomLabel = new QLabel("100%");
    m_zoomLabel->setFixedWidth(38);
    m_zoomLabel->setAlignment(Qt::AlignCenter);
    m_zoomLabel->setStyleSheet("color:#444; font-size:11px; cursor:pointer;");
    m_zoomLabel->setCursor(Qt::PointingHandCursor);
    sb->addPermanentWidget(m_zoomLabel);

    // Zoom slider
    m_zoomSlider = new QSlider(Qt::Horizontal);
    m_zoomSlider->setRange(50,200); m_zoomSlider->setValue(100);
    m_zoomSlider->setFixedWidth(90); m_zoomSlider->setFixedHeight(16);
    m_zoomSlider->setTickInterval(50);
    m_zoomSlider->setStyleSheet(
        "QSlider::groove:horizontal { height:3px; background:#c8c8c8; border-radius:2px; margin:0; }"
        "QSlider::handle:horizontal { width:14px; height:14px; margin:-6px 0; "
        "  background:white; border:2px solid #1e7145; border-radius:7px; }"
        "QSlider::sub-page:horizontal { background:#1e7145; border-radius:2px; }"
    );
    sb->addPermanentWidget(m_zoomSlider);

    // Zoom in/out buttons
    auto mkZoom = [&](const QString& t, int delta) {
        auto* b = new QToolButton;
        b->setText(t); b->setFixedSize(20,20);
        b->setStyleSheet(
            "QToolButton { border:1px solid #c8c8c8; border-radius:2px; background:#e8e8e8; "
            "font-size:12px; font-weight:bold; color:#555; }"
            "QToolButton:hover { background:#d0e8d8; border-color:#1e7145; color:#1a6b35; }"
        );
        connect(b, &QToolButton::clicked, this, [this,delta]{ setZoom(m_zoomPercent+delta); });
        return b;
    };
    sb->addPermanentWidget(mkZoom("−",-10));
    sb->addPermanentWidget(mkZoom("+",+10));

    // Connect zoom
    connect(m_zoomSlider, &QSlider::valueChanged, this, &MainWindow::onZoomChanged);
    connect(m_zoomLabel, &QLabel::linkActivated, this, [this]{
        bool ok;
        int z = QInputDialog::getInt(this,"Zoom","Zoom %:",m_zoomPercent,50,200,10,&ok);
        if (ok) setZoom(z);
    });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  RIBBON CONNECTIONS
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::connectRibbon() {
    connect(m_ribbon,&RibbonWidget::newFileRequested,  this,&MainWindow::newFile);
    connect(m_ribbon,&RibbonWidget::openFileRequested, this,&MainWindow::openFile);
    connect(m_ribbon,&RibbonWidget::saveFileRequested, this,&MainWindow::saveFile);
    connect(m_ribbon,&RibbonWidget::saveAsRequested,   this,&MainWindow::saveFileAs);

    connect(m_ribbon,&RibbonWidget::pasteRequested, this,[this]{
        QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_V,Qt::ControlModifier)); });
    connect(m_ribbon,&RibbonWidget::cutRequested,   this,[this]{
        m_core->clearCell(currentSheet(),m_view->currentRow(),m_view->currentCol()); });
    connect(m_ribbon,&RibbonWidget::copyRequested,  this,[this]{
        QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_C,Qt::ControlModifier)); });
    connect(m_ribbon,&RibbonWidget::formatPainterRequested,this,[this]{
        m_statusReady->setText("Format Painter — click a cell"); });

    connect(m_ribbon,&RibbonWidget::fontFamilyChanged, m_view,&SpreadsheetView::applyFontFamily);
    connect(m_ribbon,&RibbonWidget::fontSizeChanged,   m_view,&SpreadsheetView::applyFontSize);
    connect(m_ribbon,&RibbonWidget::boldToggled,       m_view,&SpreadsheetView::applyBold);
    connect(m_ribbon,&RibbonWidget::italicToggled,     m_view,&SpreadsheetView::applyItalic);
    connect(m_ribbon,&RibbonWidget::underlineToggled,  m_view,&SpreadsheetView::applyUnderline);
    connect(m_ribbon,&RibbonWidget::textColorChanged,  m_view,&SpreadsheetView::applyTextColor);
    connect(m_ribbon,&RibbonWidget::fillColorChanged,  m_view,&SpreadsheetView::applyFillColor);

    connect(m_ribbon,&RibbonWidget::hAlignChanged,       m_view,&SpreadsheetView::applyHAlign);
    connect(m_ribbon,&RibbonWidget::vAlignChanged,       m_view,&SpreadsheetView::applyVAlign);
    connect(m_ribbon,&RibbonWidget::wrapTextToggled,     m_view,&SpreadsheetView::applyWrapText);
    connect(m_ribbon,&RibbonWidget::mergeCellsRequested, m_view,&SpreadsheetView::mergeSelected);

    connect(m_ribbon,&RibbonWidget::numberFormatChanged,m_view,&SpreadsheetView::applyNumberFormat);
    connect(m_ribbon,&RibbonWidget::increaseDecimalRequested,this,[this]{
        auto fmt=m_view->currentCellFormat(); fmt.decimals++; m_view->applyFormat(fmt); });
    connect(m_ribbon,&RibbonWidget::decreaseDecimalRequested,this,[this]{
        auto fmt=m_view->currentCellFormat(); fmt.decimals=qMax(0,fmt.decimals-1); m_view->applyFormat(fmt); });

    connect(m_ribbon,&RibbonWidget::insertRowRequested,    m_view,&SpreadsheetView::insertRow);
    connect(m_ribbon,&RibbonWidget::deleteRowRequested,    m_view,&SpreadsheetView::deleteRow);
    connect(m_ribbon,&RibbonWidget::insertColumnRequested, m_view,&SpreadsheetView::insertColumn);
    connect(m_ribbon,&RibbonWidget::deleteColumnRequested, m_view,&SpreadsheetView::deleteColumn);
    connect(m_ribbon,&RibbonWidget::formatCellsRequested,  this,[this]{
        FormatCellsDialog dlg(m_view->currentCellFormat(),this);
        if (dlg.exec()==QDialog::Accepted) m_view->applyFormat(dlg.result()); });

    connect(m_ribbon,&RibbonWidget::autoSumRequested,   m_view,&SpreadsheetView::autoSum);
    connect(m_ribbon,&RibbonWidget::sortAscRequested,   m_view,&SpreadsheetView::sortAsc);
    connect(m_ribbon,&RibbonWidget::sortDescRequested,  m_view,&SpreadsheetView::sortDesc);
    connect(m_ribbon,&RibbonWidget::filterRequested,    this,[this]{
        m_statusReady->setText("Filter — not yet implemented"); });
    connect(m_ribbon,&RibbonWidget::findReplaceRequested,this,[this]{
        auto* dlg=new FindReplaceDialog(m_core,currentSheet(),this);
        dlg->setAttribute(Qt::WA_DeleteOnClose); dlg->show(); });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  VIEW CONNECTIONS
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::connectView() {
    connect(m_view,&SpreadsheetView::selectionFormatChanged,
            this,  &MainWindow::onSelectionChanged);
    connect(m_view->publicSelectionModel(),&QItemSelectionModel::selectionChanged,
            this,[this]{ updateSelectionStats(); });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SELECTION → FORMULA BAR + RIBBON
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::onSelectionChanged(const CellFormat& fmt, const QString& ref) {
    m_nameBox->setText(ref);

    Cell cell = m_core->getCell(currentSheet(),m_view->currentRow(),m_view->currentCol());
    m_formulaBar->blockSignals(true);
    m_formulaBar->setText(cell.formula.isEmpty() ? cell.rawValue.toString() : cell.formula);
    m_formulaBar->blockSignals(false);
    m_formulaEditing = false;
    if (m_statusReady) m_statusReady->setText("Ready");

    RibbonFormatState s;
    s.fontFamily   = fmt.font.family().isEmpty() ? "Calibri" : fmt.font.family();
    s.fontSize     = fmt.font.pointSize()>0 ? fmt.font.pointSize() : 11;
    s.bold         = fmt.bold;
    s.italic       = fmt.italic;
    s.underline    = fmt.underline;
    s.textColor    = fmt.textColor;
    s.fillColor    = fmt.fillColor;
    s.numberFormat = fmt.numberFormat;
    s.wrapText     = fmt.wrapText;
    s.hAlign       = (fmt.alignment & Qt::AlignRight) ? 2 :
                     (fmt.alignment & Qt::AlignHCenter) ? 1 : 0;
    m_ribbon->setFormatState(s);
}

void MainWindow::onFormatChanged(const CellFormat& fmt, const QString& ref) {
    onSelectionChanged(fmt, ref);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  FORMULA BAR COMMIT
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::onFormulaBarReturn() {
    QString text = m_formulaBar->text();
    int r = m_view->currentRow(), c = m_view->currentCol();
    if (r < 0 || c < 0) return;
    if (text.startsWith('='))
        m_core->setCellFormula(currentSheet(),r,c,text);
    else
        m_core->setCellValue(currentSheet(),r,c,text);
    setModified(true);
    m_formulaEditing = false;
    if (m_statusReady) m_statusReady->setText("Ready");
    m_view->setFocus();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SELECTION STATS  (Average / Count / Sum like WPS bottom bar)
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::updateSelectionStats() {
    auto sel = m_view->publicSelectedIndexes();
    if (sel.isEmpty()) { if (m_statusStats) m_statusStats->clear(); return; }

    double sum=0; int count=0; int numCount=0;
    SheetId s = currentSheet();
    for (auto& idx : sel) {
        Cell cell = m_core->getCell(s,idx.row(),idx.column());
        QVariant v = cell.cachedValue.isValid() ? cell.cachedValue : cell.rawValue;
        if (v.isValid() && !v.isNull()) {
            count++;
            bool ok; double d = v.toDouble(&ok);
            if (ok) { sum+=d; numCount++; }
        }
    }

    if (m_statusStats) {
        if (numCount > 1) {
            m_statusStats->setText(QString("Average: %1   Count: %2   Sum: %3")
                .arg(sum/numCount, 0,'g',6)
                .arg(count)
                .arg(sum, 0,'g',10));
        } else if (count == 1) {
            Cell c = m_core->getCell(s,sel.first().row(),sel.first().column());
            m_statusStats->setText((c.cachedValue.isValid() ? c.cachedValue : c.rawValue).toString());
        } else {
            m_statusStats->clear();
        }
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  ZOOM
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::onZoomChanged(int v) { setZoom(v); }

void MainWindow::setZoom(int percent) {
    m_zoomPercent = qBound(50,percent,200);
    m_zoomSlider->blockSignals(true);
    m_zoomSlider->setValue(m_zoomPercent);
    m_zoomSlider->blockSignals(false);
    m_zoomLabel->setText(QString::number(m_zoomPercent)+"%");
    m_view->setZoomFactor(m_zoomPercent/100.0);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SHEET MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::rebuildSheetTabs() {
    QSignalBlocker b(m_sheetBar);
    while (m_sheetBar->count()>0) m_sheetBar->removeTab(0);
    for (SheetId s : m_core->sheets())
        m_sheetBar->addTab(m_core->sheetName(s));
}

void MainWindow::switchSheet(int idx) {
    auto sheets = m_core->sheets();
    if (idx>=0 && idx<sheets.size()) m_view->switchSheet(sheets[idx]);
    updateSelectionStats();
}

void MainWindow::addSheet() {
    bool ok;
    QString name = QInputDialog::getText(this,"Insert Sheet","Sheet name:",
        QLineEdit::Normal,QString("Sheet%1").arg(m_core->sheets().size()+1),&ok);
    if (!ok||name.isEmpty()) return;
    SheetId s = m_core->addSheet(name);
    m_sheetBar->addTab(m_core->sheetName(s));
    m_sheetBar->setCurrentIndex(m_sheetBar->count()-1);
    setModified(true);
}

void MainWindow::removeSheet() {
    auto sheets = m_core->sheets();
    if (sheets.size()<=1) { QMessageBox::warning(this,"Delete Sheet","Cannot delete the only sheet."); return; }
    int idx = m_sheetBar->currentIndex();
    if (idx<0||idx>=sheets.size()) return;
    if (QMessageBox::question(this,"Delete Sheet",
            QString("Delete \"%1\"?").arg(m_core->sheetName(sheets[idx])),
            QMessageBox::Yes|QMessageBox::No) != QMessageBox::Yes) return;
    m_core->removeSheet(sheets[idx]);
    rebuildSheetTabs();
    m_view->switchSheet(currentSheet());
    setModified(true);
}

void MainWindow::renameSheet(int idx) {
    auto sheets = m_core->sheets();
    if (idx<0||idx>=sheets.size()) return;
    bool ok;
    QString name = QInputDialog::getText(this,"Rename Sheet","New name:",
        QLineEdit::Normal,m_core->sheetName(sheets[idx]),&ok);
    if (ok&&!name.isEmpty()) {
        m_core->renameSheet(sheets[idx],name);
        m_sheetBar->setTabText(idx,name);
        setModified(true);
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  FILE OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════════
void MainWindow::newFile() {
    if (!confirmSave()) return;
    delete m_core; m_core = createSpreadsheetCore();
    m_filePath.clear(); setModified(false);
    m_view->switchSheet(m_core->sheets().first());
    rebuildSheetTabs(); updateTitle();
    if (m_statusReady) m_statusReady->setText("Ready");
}

void MainWindow::openFile() {
    if (!confirmSave()) return;
    QString path = QFileDialog::getOpenFileName(this,"Open File",{},
        "Spreadsheet Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*)");
    if (path.isEmpty()) return;
    loadFileAsync(path);
}

void MainWindow::loadFileAsync(const QString& path) {
    m_filePath = path;
    if (m_progress) { m_progress->setVisible(true); m_progress->setRange(0,100); }
    if (m_statusReady) m_statusReady->setText("Loading "+QFileInfo(path).fileName()+"...");

    auto unused_future = QtConcurrent::run([this,path](){
        m_loader->loadChunk(path,m_core,m_core->sheets().first(),0,INT_MAX,
            [this](qint64 r,qint64 t){
                if (t>0) QMetaObject::invokeMethod(m_progress,"setValue",Qt::QueuedConnection,Q_ARG(int,(int)(r*100/t)));
            });
        QMetaObject::invokeMethod(this,[this,path](){
            if (m_progress) m_progress->setVisible(false);
            if (m_statusReady) m_statusReady->setText("Loaded: "+QFileInfo(path).fileName());
            m_view->switchSheet(m_core->sheets().first());
            rebuildSheetTabs(); setModified(false); updateTitle();
        },Qt::QueuedConnection);
    });
}

void MainWindow::saveFile() {
    if (m_filePath.isEmpty()) { saveFileAs(); return; }
    m_loader->save(m_filePath,m_core,currentSheet(),nullptr);
    setModified(false);
    if (m_statusReady) m_statusReady->setText("Saved: "+QFileInfo(m_filePath).fileName());
}

void MainWindow::saveFileAs() {
    QString path = QFileDialog::getSaveFileName(this,"Save As",m_filePath,
        "CSV Files (*.csv);;Excel Files (*.xlsx)");
    if (path.isEmpty()) return;
    m_filePath = path; saveFile();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════════════════
SheetId MainWindow::currentSheet() const {
    auto sheets = m_core->sheets();
    int idx = m_sheetBar->currentIndex();
    return (idx>=0&&idx<sheets.size()) ? sheets[idx] : sheets.first();
}
void MainWindow::setModified(bool m) { m_modified=m; updateTitle(); }
bool MainWindow::confirmSave() {
    if (!m_modified) return true;
    auto btn = QMessageBox::question(this,"Unsaved Changes","Save changes?",
        QMessageBox::Save|QMessageBox::Discard|QMessageBox::Cancel);
    if (btn==QMessageBox::Save)    { saveFile(); return true; }
    if (btn==QMessageBox::Discard) return true;
    return false;
}
void MainWindow::updateTitle() {
    QString name = m_filePath.isEmpty() ? "Untitled" : QFileInfo(m_filePath).fileName();
    setWindowTitle((m_modified?"* ":"")+name+" — OpenSheet");
}
void MainWindow::closeEvent(QCloseEvent* e) {
    if (confirmSave()) e->accept(); else e->ignore();
}
void MainWindow::resizeEvent(QResizeEvent* e) { QMainWindow::resizeEvent(e); }

QToolButton* MainWindow::makeQATButton(const QString& text, const QString& tip,
                                        const std::function<void()>& fn) {
    auto* btn = new QToolButton;
    btn->setText(text); btn->setToolTip(tip);
    btn->setFixedSize(24,24); btn->setAutoRaise(true);
    connect(btn,&QToolButton::clicked,this,fn);
    return btn;
}
