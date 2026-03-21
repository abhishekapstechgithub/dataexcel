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
#include <QSizePolicy>
#include <QStackedWidget>
#include <QPainter>
#include <QStyle>

// ── Tiny icon factory (reused from RibbonWidget pattern) ─────────────────────
static QIcon makeStatusIcon(const QColor& c, int w = 14, int h = 14) {
    QPixmap pm(w, h);
    pm.fill(Qt::transparent);
    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing);
    p.setBrush(c); p.setPen(Qt::NoPen);
    p.drawEllipse(1, 1, w-2, h-2);
    return QIcon(pm);
}

// ── Constructor ───────────────────────────────────────────────────────────────
MainWindow::MainWindow(QWidget* parent) : QMainWindow(parent) {
    setWindowTitle("OpenSheet");
    resize(1366, 820);
    setMinimumSize(800, 600);

    // ── Title bar / menu bar color (green like Excel) ────────────────────────
    menuBar()->setStyleSheet(
        "QMenuBar { background: #1e7145; color: white; font-size: 12px; padding: 2px 4px; }"
        "QMenuBar::item { padding: 4px 10px; border-radius: 3px; }"
        "QMenuBar::item:selected { background: #155835; }"
        "QMenu { background: #ffffff; color: #222; border: 1px solid #ccc; font-size: 12px; }"
        "QMenu::item { padding: 5px 28px; }"
        "QMenu::item:selected { background: #e8f4ed; color: #1a6b35; }"
        "QMenu::separator { height: 1px; background: #e0e0e0; margin: 2px 0; }"
    );

    // Load DLLs
    m_core   = createSpreadsheetCore();
    m_loader = createFileLoader();

    // ── Central widget assembly ───────────────────────────────────────────────
    auto* central = new QWidget;
    central->setObjectName("centralWidget");
    auto* vl = new QVBoxLayout(central);
    vl->setContentsMargins(0, 0, 0, 0);
    vl->setSpacing(0);

    // 1. Quick Access Toolbar (inside menu bar area)
    setupQuickAccessBar();

    // 2. Ribbon
    m_ribbon = createRibbonWidget(this);
    vl->addWidget(m_ribbon);

    // 3. Thin separator line
    auto* sep1 = new QFrame;
    sep1->setFrameShape(QFrame::HLine);
    sep1->setFixedHeight(1);
    sep1->setStyleSheet("background:#d0d0d0;");
    vl->addWidget(sep1);

    // 4. Formula bar
    setupFormulaBar();
    vl->addWidget(findChild<QWidget*>("formulaBarWidget"));

    // 5. Another separator
    auto* sep2 = new QFrame;
    sep2->setFrameShape(QFrame::HLine);
    sep2->setFixedHeight(1);
    sep2->setStyleSheet("background:#d0d0d0;");
    vl->addWidget(sep2);

    // 6. Spreadsheet view (takes all remaining space)
    m_view = new SpreadsheetView(m_core, m_core->sheets().first(), this);
    vl->addWidget(m_view, 1);

    // 7. Sheet tab bar + add button (at bottom, above status bar)
    setupSheetBar();
    vl->addWidget(findChild<QWidget*>("sheetBarWidget"));

    setCentralWidget(central);

    // 8. Setup menus and status bar
    setupStatusBar();
    connectRibbon();
    connectView();
    rebuildSheetTabs();
    updateTitle();
}

MainWindow::~MainWindow() {
    delete m_core;
    delete m_loader;
}

// ── Quick Access Toolbar ──────────────────────────────────────────────────────
void MainWindow::setupQuickAccessBar() {
    // Add toolbar items to the right side of the menu bar
    auto* mb = menuBar();

    // File menu
    auto* fileMenu = mb->addMenu("File");
    fileMenu->addAction(QIcon(), "New",      this, &MainWindow::newFile,    QKeySequence::New);
    fileMenu->addAction(QIcon(), "Open...",  this, &MainWindow::openFile,   QKeySequence::Open);
    fileMenu->addSeparator();
    fileMenu->addAction(QIcon(), "Save",     this, &MainWindow::saveFile,   QKeySequence::Save);
    fileMenu->addAction(QIcon(), "Save As...",this, &MainWindow::saveFileAs, QKeySequence::SaveAs);
    fileMenu->addSeparator();
    fileMenu->addAction(QIcon(), "Exit",     qApp, &QApplication::quit,     QKeySequence::Quit);

    // Edit menu
    auto* editMenu = mb->addMenu("Edit");
    auto* undoAct = editMenu->addAction("Undo", m_core, &ISpreadsheetCore::undo, QKeySequence::Undo);
    auto* redoAct = editMenu->addAction("Redo", m_core, &ISpreadsheetCore::redo, QKeySequence::Redo);
    editMenu->addSeparator();
    editMenu->addAction("Find && Replace...", this, [this]{
        auto* dlg = new FindReplaceDialog(m_core, currentSheet(), this);
        dlg->setAttribute(Qt::WA_DeleteOnClose); dlg->show();
    }, QKeySequence(Qt::CTRL | Qt::Key_H));

    // Sheet menu
    auto* sheetMenu = mb->addMenu("Sheet");
    sheetMenu->addAction("Add Sheet",    this, &MainWindow::addSheet);
    sheetMenu->addAction("Remove Sheet", this, &MainWindow::removeSheet);
}

// ── Formula Bar ───────────────────────────────────────────────────────────────
void MainWindow::setupFormulaBar() {
    auto* fbar = new QWidget;
    fbar->setObjectName("formulaBarWidget");
    fbar->setFixedHeight(32);
    fbar->setStyleSheet("QWidget { background: #ffffff; }");

    auto* fl = new QHBoxLayout(fbar);
    fl->setContentsMargins(6, 3, 6, 3);
    fl->setSpacing(0);

    // Name box (cell reference like A1, B3:D7)
    m_nameBox = new QLineEdit;
    m_nameBox->setObjectName("nameBox");
    m_nameBox->setFixedWidth(80);
    m_nameBox->setFixedHeight(24);
    m_nameBox->setAlignment(Qt::AlignCenter);
    m_nameBox->setStyleSheet(
        "QLineEdit { border: 1px solid #c0c0c0; border-radius: 2px; "
        "font-size: 12px; font-weight: 500; background: white; padding: 0 4px; }"
        "QLineEdit:focus { border-color: #1e7145; }"
    );
    m_nameBox->setPlaceholderText("A1");
    fl->addWidget(m_nameBox);
    fl->addSpacing(4);

    // Vertical separator
    auto* vsep = new QFrame;
    vsep->setFrameShape(QFrame::VLine);
    vsep->setFixedHeight(20);
    vsep->setStyleSheet("color: #c0c0c0;");
    fl->addWidget(vsep);
    fl->addSpacing(4);

    // fx label (clickable icon)
    m_fxLabel = new QLabel("fx");
    m_fxLabel->setFixedWidth(28);
    m_fxLabel->setFixedHeight(24);
    m_fxLabel->setAlignment(Qt::AlignCenter);
    m_fxLabel->setStyleSheet(
        "QLabel { font-size: 12px; font-style: italic; color: #1a6b35; "
        "font-weight: bold; border: 1px solid #d0d0d0; border-radius: 2px; "
        "background: #f8f8f8; }"
    );
    m_fxLabel->setCursor(Qt::PointingHandCursor);
    fl->addWidget(m_fxLabel);
    fl->addSpacing(4);

    // Another separator
    auto* vsep2 = new QFrame;
    vsep2->setFrameShape(QFrame::VLine);
    vsep2->setFixedHeight(20);
    vsep2->setStyleSheet("color: #c0c0c0;");
    fl->addWidget(vsep2);
    fl->addSpacing(4);

    // Formula / value input bar (takes remaining space)
    m_formulaBar = new QLineEdit;
    m_formulaBar->setObjectName("formulaBar");
    m_formulaBar->setFixedHeight(24);
    m_formulaBar->setPlaceholderText("Enter value or formula...");
    m_formulaBar->setStyleSheet(
        "QLineEdit { border: 1px solid #c0c0c0; border-radius: 2px; "
        "font-size: 12px; background: white; padding: 0 6px; }"
        "QLineEdit:focus { border-color: #1e7145; }"
    );
    fl->addWidget(m_formulaBar, 1);

    // Connect formula bar interactions
    connect(m_formulaBar, &QLineEdit::returnPressed, this, &MainWindow::onFormulaBarReturn);
    connect(m_formulaBar, &QLineEdit::textEdited, this, [this](const QString&){
        m_formulaEditing = true;
        m_statusReady->setText("Edit");
    });

    // Name box: navigate to cell when Enter pressed
    connect(m_nameBox, &QLineEdit::returnPressed, this, [this]{
        QString ref = m_nameBox->text().trimmed().toUpper();
        if (ref.isEmpty()) return;
        // Parse "A1" → row/col
        QRegularExpression re("^([A-Z]+)(\\d+)$");
        auto m = re.match(ref);
        if (m.hasMatch()) {
            int col = SpreadsheetTableModel::columnIndex(m.captured(1));
            int row = m.captured(2).toInt() - 1;
            m_view->setCurrentIndex(m_view->model()->index(row, col));
            m_view->setFocus();
        }
    });
}

// ── Sheet Tab Bar ─────────────────────────────────────────────────────────────
void MainWindow::setupSheetBar() {
    auto* sheetBarWidget = new QWidget;
    sheetBarWidget->setObjectName("sheetBarWidget");
    sheetBarWidget->setFixedHeight(28);
    sheetBarWidget->setStyleSheet("background: #f0f0f0; border-top: 1px solid #c0c0c0;");

    auto* hl = new QHBoxLayout(sheetBarWidget);
    hl->setContentsMargins(4, 0, 4, 0);
    hl->setSpacing(2);

    // Navigation buttons (scroll sheet tabs)
    auto makeNavBtn = [](const QString& t, const QString& tip) {
        auto* btn = new QToolButton;
        btn->setText(t);
        btn->setToolTip(tip);
        btn->setFixedSize(18, 18);
        btn->setStyleSheet("QToolButton { border: 1px solid #c0c0c0; border-radius: 2px; "
                           "background: #e8e8e8; font-size: 9px; color: #666; }"
                           "QToolButton:hover { background: #d8d8d8; }");
        return btn;
    };

    auto* btnFirst = makeNavBtn("◀◀", "First sheet");
    auto* btnPrev  = makeNavBtn("◀",  "Previous sheet");
    auto* btnNext  = makeNavBtn("▶",  "Next sheet");
    auto* btnLast  = makeNavBtn("▶▶", "Last sheet");
    hl->addWidget(btnFirst); hl->addWidget(btnPrev);
    hl->addWidget(btnNext);  hl->addWidget(btnLast);
    hl->addSpacing(4);

    // Sheet tabs
    m_sheetBar = new QTabBar;
    m_sheetBar->setTabsClosable(false);
    m_sheetBar->setMovable(true);
    m_sheetBar->setDrawBase(false);
    m_sheetBar->setExpanding(false);
    m_sheetBar->setFixedHeight(26);
    m_sheetBar->setStyleSheet(
        "QTabBar { background: transparent; }"
        "QTabBar::tab { min-width: 70px; max-width: 160px; height: 22px; "
        "  padding: 2px 12px; margin-right: 2px; font-size: 11px; "
        "  border: 1px solid #b0b0b0; border-bottom: none; border-radius: 3px 3px 0 0; "
        "  background: #e0e0e0; color: #444; }"
        "QTabBar::tab:selected { background: #ffffff; color: #1a6b35; "
        "  font-weight: 600; border-color: #b0b0b0; border-bottom: 2px solid #1e7145; }"
        "QTabBar::tab:hover:!selected { background: #ececec; }"
    );
    hl->addWidget(m_sheetBar, 1);

    // Add sheet "+" button
    m_addSheetBtn = new QToolButton;
    m_addSheetBtn->setText("+");
    m_addSheetBtn->setToolTip("Add sheet");
    m_addSheetBtn->setFixedSize(22, 22);
    m_addSheetBtn->setStyleSheet(
        "QToolButton { border: 1px solid #c0c0c0; border-radius: 11px; "
        "background: #e8e8e8; font-size: 14px; color: #555; font-weight: bold; }"
        "QToolButton:hover { background: #d0d0d0; color: #1a6b35; }"
    );
    hl->addWidget(m_addSheetBtn);

    // Connections
    connect(m_sheetBar,   &QTabBar::currentChanged,     this, &MainWindow::switchSheet);
    connect(m_sheetBar,   &QTabBar::tabBarDoubleClicked,this, &MainWindow::renameSheet);
    connect(m_addSheetBtn,&QToolButton::clicked,         this, &MainWindow::addSheet);

    // Nav button connections
    connect(btnFirst, &QToolButton::clicked, this, [this]{ m_sheetBar->setCurrentIndex(0); });
    connect(btnLast,  &QToolButton::clicked, this, [this]{ m_sheetBar->setCurrentIndex(m_sheetBar->count()-1); });
    connect(btnPrev,  &QToolButton::clicked, this, [this]{
        m_sheetBar->setCurrentIndex(qMax(0, m_sheetBar->currentIndex()-1)); });
    connect(btnNext,  &QToolButton::clicked, this, [this]{
        m_sheetBar->setCurrentIndex(qMin(m_sheetBar->count()-1, m_sheetBar->currentIndex()+1)); });

    // Right-click context menu on tabs
    m_sheetBar->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_sheetBar, &QWidget::customContextMenuRequested, this, [this](const QPoint& pos){
        int idx = m_sheetBar->tabAt(pos);
        if (idx < 0) return;
        QMenu menu;
        menu.addAction("Rename",  [this,idx]{ renameSheet(idx); });
        menu.addAction("Delete",  [this]{ removeSheet(); });
        menu.addSeparator();
        menu.addAction("Insert Sheet...", [this]{ addSheet(); });
        menu.exec(m_sheetBar->mapToGlobal(pos));
    });
}

// ── Status Bar ────────────────────────────────────────────────────────────────
void MainWindow::setupStatusBar() {
    auto* sb = statusBar();
    sb->setStyleSheet(
        "QStatusBar { background: #f0f0f0; border-top: 1px solid #c0c0c0; font-size: 11px; }"
        "QStatusBar::item { border: none; }"
    );

    // Left: Ready indicator with green dot
    m_statusReady = new QLabel("Ready");
    m_statusReady->setStyleSheet("color: #444; padding: 0 8px;");
    sb->addWidget(m_statusReady);

    // Separator
    auto* sep1 = new QFrame; sep1->setFrameShape(QFrame::VLine);
    sep1->setStyleSheet("color: #c0c0c0;"); sep1->setFixedHeight(14);
    sb->addWidget(sep1);

    // Cell mode indicator
    m_statusMode = new QLabel("Ready");
    m_statusMode->setStyleSheet("color: #666; padding: 0 8px;");
    sb->addWidget(m_statusMode);

    // Progress bar (hidden by default)
    m_progress = new QProgressBar;
    m_progress->setFixedWidth(180);
    m_progress->setFixedHeight(14);
    m_progress->setVisible(false);
    m_progress->setStyleSheet(
        "QProgressBar { border: 1px solid #c0c0c0; border-radius: 2px; background: #e8e8e8; }"
        "QProgressBar::chunk { background: #1e7145; border-radius: 1px; }"
    );
    sb->addWidget(m_progress);

    // Right side: selection stats
    m_statusStats = new QLabel;
    m_statusStats->setStyleSheet("color: #444; padding: 0 12px;");
    m_statusStats->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
    sb->addPermanentWidget(m_statusStats);

    // Separator before zoom
    auto* sep2 = new QFrame; sep2->setFrameShape(QFrame::VLine);
    sep2->setStyleSheet("color: #c0c0c0;"); sep2->setFixedHeight(14);
    sb->addPermanentWidget(sep2);

    // View mode buttons (Normal / Page Layout / Page Break)
    auto makeViewBtn = [&](const QString& tip) {
        auto* btn = new QToolButton;
        btn->setToolTip(tip);
        btn->setFixedSize(18, 18);
        btn->setCheckable(true);
        btn->setStyleSheet(
            "QToolButton { border: 1px solid transparent; border-radius: 2px; background: transparent; }"
            "QToolButton:hover { border-color: #c0c0c0; background: #e0e0e0; }"
            "QToolButton:checked { background: #c8e6d0; border-color: #2d7d46; }"
        );
        return btn;
    };
    auto* btnNormal = makeViewBtn("Normal view");
    auto* btnPage   = makeViewBtn("Page layout view");
    auto* btnBreak  = makeViewBtn("Page break view");
    // Draw simple icons using text
    btnNormal->setText("▦"); btnPage->setText("▤"); btnBreak->setText("▨");
    btnNormal->setChecked(true);
    sb->addPermanentWidget(btnNormal);
    sb->addPermanentWidget(btnPage);
    sb->addPermanentWidget(btnBreak);

    // Separator
    auto* sep3 = new QFrame; sep3->setFrameShape(QFrame::VLine);
    sep3->setStyleSheet("color: #c0c0c0;"); sep3->setFixedHeight(14);
    sb->addPermanentWidget(sep3);

    // Zoom out button
    auto* zoomOut = new QToolButton;
    zoomOut->setText("−");
    zoomOut->setFixedSize(18, 18);
    zoomOut->setStyleSheet("QToolButton { border: 1px solid #c0c0c0; border-radius: 2px; "
                           "background: #e8e8e8; font-size: 12px; }"
                           "QToolButton:hover { background: #d0d0d0; }");
    sb->addPermanentWidget(zoomOut);

    // Zoom slider
    m_zoomSlider = new QSlider(Qt::Horizontal);
    m_zoomSlider->setRange(50, 200);
    m_zoomSlider->setValue(100);
    m_zoomSlider->setFixedWidth(80);
    m_zoomSlider->setFixedHeight(16);
    m_zoomSlider->setTickInterval(50);
    m_zoomSlider->setStyleSheet(
        "QSlider::groove:horizontal { height: 3px; background: #c0c0c0; border-radius: 1px; }"
        "QSlider::handle:horizontal { width: 12px; height: 12px; margin: -5px 0; "
        "  background: #1e7145; border-radius: 6px; }"
        "QSlider::sub-page:horizontal { background: #1e7145; border-radius: 1px; }"
    );
    sb->addPermanentWidget(m_zoomSlider);

    // Zoom in button
    auto* zoomIn = new QToolButton;
    zoomIn->setText("+");
    zoomIn->setFixedSize(18, 18);
    zoomIn->setStyleSheet(zoomOut->styleSheet());
    sb->addPermanentWidget(zoomIn);

    // Zoom percentage label
    m_zoomLabel = new QLabel("100%");
    m_zoomLabel->setFixedWidth(36);
    m_zoomLabel->setAlignment(Qt::AlignCenter);
    m_zoomLabel->setStyleSheet("color: #444; font-size: 11px;");
    m_zoomLabel->setCursor(Qt::PointingHandCursor);
    sb->addPermanentWidget(m_zoomLabel);

    // Connect zoom
    connect(m_zoomSlider, &QSlider::valueChanged, this, &MainWindow::onZoomChanged);
    connect(zoomOut, &QToolButton::clicked, this, [this]{
        setZoom(qMax(50, m_zoomPercent - 10)); });
    connect(zoomIn,  &QToolButton::clicked, this, [this]{
        setZoom(qMin(200, m_zoomPercent + 10)); });
    connect(m_zoomLabel, &QLabel::linkActivated, this, [this]{
        bool ok;
        int z = QInputDialog::getInt(this, "Zoom", "Zoom %:", m_zoomPercent, 50, 200, 10, &ok);
        if (ok) setZoom(z);
    });
    // Make zoom label clickable
    m_zoomLabel->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_zoomLabel, &QLabel::customContextMenuRequested, this, [this](const QPoint&){
        bool ok;
        int z = QInputDialog::getInt(this, "Zoom", "Zoom %:", m_zoomPercent, 50, 200, 10, &ok);
        if (ok) setZoom(z);
    });
}

// ── Connect Ribbon ────────────────────────────────────────────────────────────
void MainWindow::connectRibbon() {
    // File
    connect(m_ribbon, &RibbonWidget::newFileRequested,  this, &MainWindow::newFile);
    connect(m_ribbon, &RibbonWidget::openFileRequested, this, &MainWindow::openFile);
    connect(m_ribbon, &RibbonWidget::saveFileRequested, this, &MainWindow::saveFile);
    connect(m_ribbon, &RibbonWidget::saveAsRequested,   this, &MainWindow::saveFileAs);

    // Clipboard
    connect(m_ribbon, &RibbonWidget::pasteRequested, this, [this]{
        QApplication::sendEvent(m_view,
            new QKeyEvent(QEvent::KeyPress, Qt::Key_V, Qt::ControlModifier)); });
    connect(m_ribbon, &RibbonWidget::cutRequested, this, [this]{
        m_core->clearCell(currentSheet(), m_view->currentRow(), m_view->currentCol()); });
    connect(m_ribbon, &RibbonWidget::copyRequested, this, [this]{
        QApplication::sendEvent(m_view,
            new QKeyEvent(QEvent::KeyPress, Qt::Key_C, Qt::ControlModifier)); });
    connect(m_ribbon, &RibbonWidget::formatPainterRequested, this, [this]{
        m_statusReady->setText("Format Painter active — click a cell to apply"); });

    // Font
    connect(m_ribbon, &RibbonWidget::fontFamilyChanged, m_view, &SpreadsheetView::applyFontFamily);
    connect(m_ribbon, &RibbonWidget::fontSizeChanged,   m_view, &SpreadsheetView::applyFontSize);
    connect(m_ribbon, &RibbonWidget::boldToggled,       m_view, &SpreadsheetView::applyBold);
    connect(m_ribbon, &RibbonWidget::italicToggled,     m_view, &SpreadsheetView::applyItalic);
    connect(m_ribbon, &RibbonWidget::underlineToggled,  m_view, &SpreadsheetView::applyUnderline);
    connect(m_ribbon, &RibbonWidget::textColorChanged,  m_view, &SpreadsheetView::applyTextColor);
    connect(m_ribbon, &RibbonWidget::fillColorChanged,  m_view, &SpreadsheetView::applyFillColor);

    // Alignment
    connect(m_ribbon, &RibbonWidget::hAlignChanged,      m_view, &SpreadsheetView::applyHAlign);
    connect(m_ribbon, &RibbonWidget::vAlignChanged,      m_view, &SpreadsheetView::applyVAlign);
    connect(m_ribbon, &RibbonWidget::wrapTextToggled,    m_view, &SpreadsheetView::applyWrapText);
    connect(m_ribbon, &RibbonWidget::mergeCellsRequested,m_view, &SpreadsheetView::mergeSelected);

    // Number
    connect(m_ribbon, &RibbonWidget::numberFormatChanged,m_view, &SpreadsheetView::applyNumberFormat);
    connect(m_ribbon, &RibbonWidget::increaseDecimalRequested, this, [this]{
        auto fmt = m_view->currentCellFormat(); fmt.decimals++; m_view->applyFormat(fmt); });
    connect(m_ribbon, &RibbonWidget::decreaseDecimalRequested, this, [this]{
        auto fmt = m_view->currentCellFormat();
        fmt.decimals = qMax(0, fmt.decimals-1); m_view->applyFormat(fmt); });

    // Cells
    connect(m_ribbon, &RibbonWidget::insertRowRequested,    m_view, &SpreadsheetView::insertRow);
    connect(m_ribbon, &RibbonWidget::deleteRowRequested,    m_view, &SpreadsheetView::deleteRow);
    connect(m_ribbon, &RibbonWidget::insertColumnRequested, m_view, &SpreadsheetView::insertColumn);
    connect(m_ribbon, &RibbonWidget::deleteColumnRequested, m_view, &SpreadsheetView::deleteColumn);
    connect(m_ribbon, &RibbonWidget::formatCellsRequested,  this, [this]{
        FormatCellsDialog dlg(m_view->currentCellFormat(), this);
        if (dlg.exec() == QDialog::Accepted) m_view->applyFormat(dlg.result()); });

    // Editing
    connect(m_ribbon, &RibbonWidget::autoSumRequested,    m_view, &SpreadsheetView::autoSum);
    connect(m_ribbon, &RibbonWidget::sortAscRequested,    m_view, &SpreadsheetView::sortAsc);
    connect(m_ribbon, &RibbonWidget::sortDescRequested,   m_view, &SpreadsheetView::sortDesc);
    connect(m_ribbon, &RibbonWidget::filterRequested, this, [this]{
        m_statusReady->setText("Filter — not yet implemented"); });
    connect(m_ribbon, &RibbonWidget::findReplaceRequested, this, [this]{
        auto* dlg = new FindReplaceDialog(m_core, currentSheet(), this);
        dlg->setAttribute(Qt::WA_DeleteOnClose); dlg->show(); });
}

// ── Connect View ──────────────────────────────────────────────────────────────
void MainWindow::connectView() {
    // Selection → formula bar + ribbon state
    connect(m_view, &SpreadsheetView::selectionFormatChanged,
            this,   &MainWindow::onSelectionChanged);

    // Selection → status bar stats
    connect(m_view->selectionModel(), &QItemSelectionModel::selectionChanged,
            this, [this]{ updateSelectionStats(); });
}

// ── Selection Changed ─────────────────────────────────────────────────────────
void MainWindow::onSelectionChanged(const CellFormat& fmt, const QString& ref) {
    // Update name box
    m_nameBox->setText(ref);

    // Update formula bar with cell content
    Cell cell = m_core->getCell(currentSheet(),
                                m_view->currentRow(), m_view->currentCol());
    m_formulaBar->blockSignals(true);
    m_formulaBar->setText(cell.formula.isEmpty()
        ? cell.rawValue.toString() : cell.formula);
    m_formulaBar->blockSignals(false);
    m_formulaEditing = false;
    m_statusReady->setText("Ready");

    // Update ribbon format state
    RibbonFormatState state;
    state.fontFamily   = fmt.font.family().isEmpty() ? "Calibri" : fmt.font.family();
    state.fontSize     = fmt.font.pointSize() > 0 ? fmt.font.pointSize() : 11;
    state.bold         = fmt.bold;
    state.italic       = fmt.italic;
    state.underline    = fmt.underline;
    state.textColor    = fmt.textColor;
    state.fillColor    = fmt.fillColor;
    state.numberFormat = fmt.numberFormat;
    state.wrapText     = fmt.wrapText;
    state.hAlign       = (fmt.alignment & Qt::AlignRight) ? 2 :
                         (fmt.alignment & Qt::AlignHCenter) ? 1 : 0;
    m_ribbon->setFormatState(state);
}

// ── Formula Bar Return ────────────────────────────────────────────────────────
void MainWindow::onFormulaBarReturn() {
    QString text = m_formulaBar->text();
    int r = m_view->currentRow(), c = m_view->currentCol();
    if (r < 0 || c < 0) return;
    if (text.startsWith('='))
        m_core->setCellFormula(currentSheet(), r, c, text);
    else
        m_core->setCellValue(currentSheet(), r, c, text);
    setModified(true);
    m_formulaEditing = false;
    m_statusReady->setText("Ready");
    m_view->setFocus();
}

// ── Update Selection Stats ─────────────────────────────────────────────────────
void MainWindow::updateSelectionStats() {
    auto selected = m_view->selectedIndexes();
    if (selected.isEmpty()) { m_statusStats->clear(); return; }

    double sum = 0; int count = 0; int numCount = 0;
    SheetId s = currentSheet();
    for (auto& idx : selected) {
        Cell cell = m_core->getCell(s, idx.row(), idx.column());
        QVariant v = cell.cachedValue.isValid() ? cell.cachedValue : cell.rawValue;
        if (v.isValid() && !v.isNull()) {
            count++;
            bool ok;
            double d = v.toDouble(&ok);
            if (ok) { sum += d; numCount++; }
        }
    }

    QString stats;
    if (numCount > 1) {
        stats = QString("Average: %1   Count: %2   Sum: %3")
            .arg(numCount > 0 ? sum/numCount : 0, 0, 'g', 6)
            .arg(count)
            .arg(sum, 0, 'g', 10);
    } else if (count == 1) {
        Cell cell = m_core->getCell(s,
            selected.first().row(), selected.first().column());
        stats = cell.cachedValue.isValid()
            ? cell.cachedValue.toString() : cell.rawValue.toString();
    }
    m_statusStats->setText(stats);
}

// ── Zoom ──────────────────────────────────────────────────────────────────────
void MainWindow::onZoomChanged(int value) {
    setZoom(value);
}

void MainWindow::setZoom(int percent) {
    m_zoomPercent = qBound(50, percent, 200);
    m_zoomSlider->blockSignals(true);
    m_zoomSlider->setValue(m_zoomPercent);
    m_zoomSlider->blockSignals(false);
    m_zoomLabel->setText(QString::number(m_zoomPercent) + "%");

    // Scale the view font
    qreal scale = m_zoomPercent / 100.0;
    m_view->setZoomFactor(scale);
}

// ── Sheet Management ──────────────────────────────────────────────────────────
void MainWindow::rebuildSheetTabs() {
    QSignalBlocker b(m_sheetBar);
    while (m_sheetBar->count() > 0) m_sheetBar->removeTab(0);
    for (SheetId s : m_core->sheets())
        m_sheetBar->addTab(m_core->sheetName(s));
}

void MainWindow::switchSheet(int idx) {
    auto sheets = m_core->sheets();
    if (idx >= 0 && idx < sheets.size()) {
        m_view->switchSheet(sheets[idx]);
        updateSelectionStats();
    }
}

void MainWindow::addSheet() {
    bool ok;
    QString name = QInputDialog::getText(this, "New Sheet", "Sheet name:",
        QLineEdit::Normal, QString("Sheet%1").arg(m_core->sheets().size()+1), &ok);
    if (!ok || name.isEmpty()) return;
    SheetId s = m_core->addSheet(name);
    m_sheetBar->addTab(m_core->sheetName(s));
    m_sheetBar->setCurrentIndex(m_sheetBar->count()-1);
    setModified(true);
}

void MainWindow::removeSheet() {
    auto sheets = m_core->sheets();
    if (sheets.size() <= 1) {
        QMessageBox::warning(this, "Remove Sheet", "Cannot remove the last sheet."); return;
    }
    int idx = m_sheetBar->currentIndex();
    if (idx < 0 || idx >= sheets.size()) return;
    if (QMessageBox::question(this, "Remove Sheet",
            QString("Delete \"%1\"?").arg(m_core->sheetName(sheets[idx])),
            QMessageBox::Yes|QMessageBox::No) != QMessageBox::Yes) return;
    m_core->removeSheet(sheets[idx]);
    rebuildSheetTabs();
    m_view->switchSheet(currentSheet());
    setModified(true);
}

void MainWindow::renameSheet(int idx) {
    auto sheets = m_core->sheets();
    if (idx < 0 || idx >= sheets.size()) return;
    bool ok;
    QString name = QInputDialog::getText(this, "Rename Sheet", "New name:",
        QLineEdit::Normal, m_core->sheetName(sheets[idx]), &ok);
    if (ok && !name.isEmpty()) {
        m_core->renameSheet(sheets[idx], name);
        m_sheetBar->setTabText(idx, name);
        setModified(true);
    }
}

// ── File operations ───────────────────────────────────────────────────────────
void MainWindow::newFile() {
    if (!confirmSave()) return;
    delete m_core;
    m_core = createSpreadsheetCore();
    m_filePath.clear();
    setModified(false);
    m_view->switchSheet(m_core->sheets().first());
    rebuildSheetTabs();
    updateTitle();
    m_statusReady->setText("Ready");
}

void MainWindow::openFile() {
    if (!confirmSave()) return;
    QString path = QFileDialog::getOpenFileName(this, "Open File", {},
        "Spreadsheet Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*)");
    if (path.isEmpty()) return;
    loadFileAsync(path);
}

void MainWindow::loadFileAsync(const QString& path) {
    m_filePath = path;
    m_progress->setVisible(true);
    m_progress->setRange(0, 100);
    m_statusReady->setText("Loading " + QFileInfo(path).fileName() + "...");

    auto unused_future = QtConcurrent::run([this, path]() {
        m_loader->loadChunk(path, m_core, m_core->sheets().first(), 0, INT_MAX,
            [this](qint64 read, qint64 total) {
                if (total > 0)
                    QMetaObject::invokeMethod(m_progress, "setValue",
                        Qt::QueuedConnection,
                        Q_ARG(int, (int)(read * 100 / total)));
            });
        QMetaObject::invokeMethod(this, [this, path]() {
            m_progress->setVisible(false);
            m_statusReady->setText("Loaded: " + QFileInfo(path).fileName());
            m_view->switchSheet(m_core->sheets().first());
            rebuildSheetTabs();
            setModified(false);
            updateTitle();
        }, Qt::QueuedConnection);
    });
}

void MainWindow::saveFile() {
    if (m_filePath.isEmpty()) { saveFileAs(); return; }
    m_loader->save(m_filePath, m_core, currentSheet(), nullptr);
    setModified(false);
    m_statusReady->setText("Saved: " + QFileInfo(m_filePath).fileName());
}

void MainWindow::saveFileAs() {
    QString path = QFileDialog::getSaveFileName(this, "Save As", m_filePath,
        "CSV Files (*.csv);;Excel Files (*.xlsx)");
    if (path.isEmpty()) return;
    m_filePath = path;
    saveFile();
}

// ── Helpers ───────────────────────────────────────────────────────────────────
SheetId MainWindow::currentSheet() const {
    auto sheets = m_core->sheets();
    int idx = m_sheetBar->currentIndex();
    return (idx >= 0 && idx < sheets.size()) ? sheets[idx] : sheets.first();
}

void MainWindow::setModified(bool m) {
    m_modified = m;
    updateTitle();
}

bool MainWindow::confirmSave() {
    if (!m_modified) return true;
    auto btn = QMessageBox::question(this, "Unsaved Changes",
        "Save changes before continuing?",
        QMessageBox::Save | QMessageBox::Discard | QMessageBox::Cancel);
    if (btn == QMessageBox::Save)    { saveFile(); return true; }
    if (btn == QMessageBox::Discard) return true;
    return false;
}

void MainWindow::updateTitle() {
    QString name = m_filePath.isEmpty() ? "Untitled" : QFileInfo(m_filePath).fileName();
    setWindowTitle((m_modified ? "* " : "") + name + " — OpenSheet");
}

void MainWindow::onFormatChanged(const CellFormat& fmt, const QString& ref) {
    onSelectionChanged(fmt, ref);
}

void MainWindow::closeEvent(QCloseEvent* e) {
    if (confirmSave()) e->accept();
    else e->ignore();
}

void MainWindow::resizeEvent(QResizeEvent* e) {
    QMainWindow::resizeEvent(e);
}

QToolButton* MainWindow::makeQATButton(const QString& text, const QString& tip,
                                        const std::function<void()>& fn)
{
    auto* btn = new QToolButton;
    btn->setText(text);
    btn->setToolTip(tip);
    btn->setFixedSize(24, 24);
    btn->setAutoRaise(true);
    connect(btn, &QToolButton::clicked, this, fn);
    return btn;
}
