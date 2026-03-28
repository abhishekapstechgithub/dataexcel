// ═══════════════════════════════════════════════════════════════════════════════
//  MainWindow.cpp — Main Excel-like application window
// ═══════════════════════════════════════════════════════════════════════════════
#include "MainWindow.h"
#include "FormulaBar.h"
#include "SpreadsheetView.h"
#include "SheetTabBar.h"
#include "FindReplaceDialog.h"
#include "FormatCellsDialog.h"
#include <SpreadsheetCore/ISpreadsheetCore.h>
#include <FileLoader/IFileLoader.h>
#include <RibbonUI/RibbonWidget.h>
#include <PluginSystem/IPlugin.h>

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QWidget>
#include <QSplitter>
#include <QStatusBar>
#include <QMenuBar>
#include <QLabel>
#include <QSlider>
#include <QToolButton>
#include <QFileDialog>
#include <QMessageBox>
#include <QInputDialog>
#include <QProgressDialog>
#include <QTimer>
#include <QCloseEvent>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QMimeData>
#include <QSettings>
#include <QApplication>
#include <QLibrary>
#include <QThread>
#include <QFuture>
#include <QtConcurrent/QtConcurrent>

// ── DLL factory typedefs ──────────────────────────────────────────────────────
using CreateCoreFunc          = ISpreadsheetCore* (*)();
using CreateRibbonFunc        = RibbonWidget*     (*)(QWidget*);
using CreateLoaderFunc        = IFileLoader*      (*)();
using CreatePluginManagerFunc = PluginManager*    (*)();

// ════════════════════════════════════════════════════════════════════════════
MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
{
    setWindowTitle("OpenSheet");
    setMinimumSize(1024, 600);
    resize(1366, 768);
    setAcceptDrops(true);

    // ── Load SpreadsheetCore DLL ──────────────────────────────────────────────
    QLibrary coreLib("SpreadsheetCore");
    if (coreLib.load()) {
        auto fn = (CreateCoreFunc)coreLib.resolve("createSpreadsheetCore");
        if (fn) m_core = fn();
    }

    // Fallback: try direct link
    if (!m_core) {
        m_core = createSpreadsheetCore();
    }

    if (m_core) {
        // Create initial sheet
        m_core->addSheet("Sheet1");
    }

    setupRibbon();
    setupFormulaBar();
    setupView();
    setupSheetTabs();
    setupStatusBar();
    setupActions();
    updateTitle();

    // ── Load plugins (from plugins/ subdirectory next to executable) ──────────
    {
        QLibrary psysLib("PluginSystem");
        if (psysLib.load()) {
            auto fn = (CreatePluginManagerFunc)psysLib.resolve("createPluginManager");
            if (fn) m_pluginManager = fn();
        }
        if (m_pluginManager) {
            m_pluginManager->setCore(m_core);
            QString pluginDir = QApplication::applicationDirPath() + "/plugins";
            m_pluginManager->loadPluginsFromDirectory(pluginDir);
        }
    }

    // ── Restore window geometry ───────────────────────────────────────────────
    QSettings s("OpenSheet", "OpenSheet");
    restoreGeometry(s.value("mainWindowGeometry").toByteArray());
    restoreState(s.value("mainWindowState").toByteArray());
}

MainWindow::~MainWindow()
{
    if (m_pluginManager) {
        m_pluginManager->unloadAll();
        delete m_pluginManager;
        m_pluginManager = nullptr;
    }
    if (m_core) {
        delete m_core;
        m_core = nullptr;
    }
}

// ── Setup ─────────────────────────────────────────────────────────────────────
void MainWindow::setupRibbon()
{
    // Load RibbonUI DLL
    QLibrary ribbonLib("RibbonUI");
    if (ribbonLib.load()) {
        auto fn = (CreateRibbonFunc)ribbonLib.resolve("createRibbonWidget");
        if (fn) m_ribbon = fn(this);
    }
    if (!m_ribbon) {
        m_ribbon = createRibbonWidget(this);
    }
    if (!m_ribbon) return;

    // ribbon is embedded in the central widget layout — no QToolBar wrapper needed

    // ── Connect ribbon signals ─────────────────────────────────────────────────
    connect(m_ribbon, &RibbonWidget::newFileRequested,     this, &MainWindow::onNew);
    connect(m_ribbon, &RibbonWidget::openFileRequested,    this, &MainWindow::onOpen);
    connect(m_ribbon, &RibbonWidget::saveFileRequested,    this, &MainWindow::onSave);
    connect(m_ribbon, &RibbonWidget::saveAsRequested,      this, &MainWindow::onSaveAs);
    connect(m_ribbon, &RibbonWidget::copyRequested,        this, &MainWindow::onCopy);
    connect(m_ribbon, &RibbonWidget::cutRequested,         this, &MainWindow::onCut);
    connect(m_ribbon, &RibbonWidget::pasteRequested,       this, &MainWindow::onPaste);
    connect(m_ribbon, &RibbonWidget::boldToggled,          this, &MainWindow::onBoldToggled);
    connect(m_ribbon, &RibbonWidget::italicToggled,        this, &MainWindow::onItalicToggled);
    connect(m_ribbon, &RibbonWidget::underlineToggled,     this, &MainWindow::onUnderlineToggled);
    connect(m_ribbon, &RibbonWidget::fontFamilyChanged,    this, &MainWindow::onFontFamilyChanged);
    connect(m_ribbon, &RibbonWidget::fontSizeChanged,      this, &MainWindow::onFontSizeChanged);
    connect(m_ribbon, &RibbonWidget::textColorChanged,     this, &MainWindow::onTextColorChanged);
    connect(m_ribbon, &RibbonWidget::fillColorChanged,     this, &MainWindow::onFillColorChanged);
    connect(m_ribbon, &RibbonWidget::hAlignChanged,        this, &MainWindow::onHAlignChanged);
    connect(m_ribbon, &RibbonWidget::vAlignChanged,        this, &MainWindow::onVAlignChanged);
    connect(m_ribbon, &RibbonWidget::wrapTextToggled,      this, &MainWindow::onWrapTextToggled);
    connect(m_ribbon, &RibbonWidget::mergeCellsRequested,  this, &MainWindow::onMergeCells);
    connect(m_ribbon, &RibbonWidget::numberFormatChanged,  this, &MainWindow::onNumberFormatChanged);
    connect(m_ribbon, &RibbonWidget::insertRowRequested,   this, &MainWindow::onInsertRow);
    connect(m_ribbon, &RibbonWidget::deleteRowRequested,   this, &MainWindow::onDeleteRow);
    connect(m_ribbon, &RibbonWidget::insertColumnRequested,this, &MainWindow::onInsertColumn);
    connect(m_ribbon, &RibbonWidget::deleteColumnRequested,this, &MainWindow::onDeleteColumn);
    connect(m_ribbon, &RibbonWidget::autoSumRequested,     this, &MainWindow::onAutoSum);
    connect(m_ribbon, &RibbonWidget::sortAscRequested,     this, &MainWindow::onSortAsc);
    connect(m_ribbon, &RibbonWidget::sortDescRequested,    this, &MainWindow::onSortDesc);
    connect(m_ribbon, &RibbonWidget::filterRequested,      this, &MainWindow::onFilter);
    connect(m_ribbon, &RibbonWidget::findReplaceRequested,  this, &MainWindow::onFindReplace);
    connect(m_ribbon, &RibbonWidget::formatCellsRequested,  this, &MainWindow::onFormatCells);
    connect(m_ribbon, &RibbonWidget::insertChartRequested,  this,
            [this](const QString& type){ onInsertChart(); Q_UNUSED(type) });

    // ── New signals ───────────────────────────────────────────────────────────
    connect(m_ribbon, &RibbonWidget::findReplaceRequested, this, &MainWindow::onFindReplace);
    connect(m_ribbon, &RibbonWidget::printPreviewRequested,this, [this]{ /* TODO: print preview */ });
    connect(m_ribbon, &RibbonWidget::marginsRequested,     this, [this]{ /* TODO: margins dialog */ });
    connect(m_ribbon, &RibbonWidget::orientationChanged,   this, [this](bool){ /* TODO: orientation */ });
    connect(m_ribbon, &RibbonWidget::spellCheckRequested,  this, [this]{ /* TODO: spell check */ });
    connect(m_ribbon, &RibbonWidget::newCommentRequested,  this, [this]{ /* TODO: new comment */ });
    connect(m_ribbon, &RibbonWidget::protectSheetRequested,   this, [this]{ /* TODO: protect sheet */ });
    connect(m_ribbon, &RibbonWidget::protectWorkbookRequested,this, [this]{ /* TODO: protect workbook */ });
    connect(m_ribbon, &RibbonWidget::showGridlinesToggled, this, [this](bool on){ Q_UNUSED(on); m_view->refresh(); });
    connect(m_ribbon, &RibbonWidget::showHeadingsToggled,  this, [this](bool){ m_view->refresh(); });
    connect(m_ribbon, &RibbonWidget::showFormulaBarToggled,this, [this](bool on){
        if (m_formulaBar) m_formulaBar->setVisible(on);
    });
    connect(m_ribbon, &RibbonWidget::zoomToValueRequested, this, [this](int pct){
        m_view->setZoom(pct);
        if (m_zoomSlider) m_zoomSlider->setValue(pct);
        if (m_zoomLabel)  m_zoomLabel->setText(QString("%1%").arg(pct));
    });
    connect(m_ribbon, &RibbonWidget::freezePanesRequested, this, [this](const QString& mode){
        if      (mode == "row")      m_view->setFreezeRow(1);
        else if (mode == "col")      m_view->setFreezeCol(1);
        else if (mode == "unfreeze"){ m_view->setFreezeRow(0); m_view->setFreezeCol(0); }
        else                         m_view->setFreezeRow(m_view->currentRow());
    });
}

void MainWindow::setupFormulaBar()
{
    m_formulaBar = new FormulaBar(this);
    connect(m_formulaBar, &FormulaBar::formulaCommitted,
            this, &MainWindow::onFormulaCommitted);
    connect(m_formulaBar, &FormulaBar::nameBoxNavigate,
            this, &MainWindow::onNameBoxNavigate);
    connect(m_formulaBar, &FormulaBar::editCancelled,
            this, [this]{ if (m_view) m_view->cancelEdit(); });
}

void MainWindow::setupView()
{
    m_view = new SpreadsheetView(this);
    if (m_core) m_view->setCore(m_core);

    connect(m_view, &SpreadsheetView::currentCellChanged,
            this, &MainWindow::onCellChanged);
    connect(m_view, &SpreadsheetView::cellEdited, this,
            [this](int r, int c, const QString&) {
                m_modified = true;
                updateTitle();
                onCellChanged(r, c);
            });
    connect(m_view, &SpreadsheetView::contextMenuRequested, this,
            [this](const QPoint& gp, int r, int c) {
                QMenu menu(this);
                menu.addAction("Cut",   this, &MainWindow::onCut);
                menu.addAction("Copy",  this, &MainWindow::onCopy);
                menu.addAction("Paste", this, &MainWindow::onPaste);
                menu.addSeparator();
                menu.addAction("Format Cells...", this, &MainWindow::onFormatCells);
                menu.addSeparator();
                menu.addAction("Insert Row",    this, &MainWindow::onInsertRow);
                menu.addAction("Delete Row",    this, &MainWindow::onDeleteRow);
                menu.addAction("Insert Column", this, &MainWindow::onInsertColumn);
                menu.addAction("Delete Column", this, &MainWindow::onDeleteColumn);
                menu.exec(gp);
            });

    // ── Assemble central widget ────────────────────────────────────────────────
    auto* central = new QWidget(this);
    auto* vbox    = new QVBoxLayout(central);
    vbox->setSpacing(0);
    vbox->setContentsMargins(0, 0, 0, 0);

    if (m_ribbon)     vbox->addWidget(m_ribbon);
    if (m_formulaBar) vbox->addWidget(m_formulaBar);
    vbox->addWidget(m_view, 1);

    setCentralWidget(central);
}

void MainWindow::setupSheetTabs()
{
    m_sheetBar = new SheetTabBar(this);
    m_sheetBar->addTab("Sheet1");

    // Insert sheetBar into the central widget layout
    auto* vbox = qobject_cast<QVBoxLayout*>(centralWidget()->layout());
    if (vbox) vbox->addWidget(m_sheetBar);

    connect(m_sheetBar, &SheetTabBar::tabActivated,
            this, &MainWindow::onSheetTabActivated);
    connect(m_sheetBar, &SheetTabBar::addSheetRequested,
            this, &MainWindow::onAddSheet);
    connect(m_sheetBar, &SheetTabBar::renameSheetRequested,
            this, &MainWindow::onRenameSheet);
    connect(m_sheetBar, &SheetTabBar::deleteSheetRequested,
            this, &MainWindow::onDeleteSheet);
}

void MainWindow::setupStatusBar()
{
    auto* sb = statusBar();
    sb->setStyleSheet("QStatusBar { background: #F2F2F2; border-top: 1px solid #D0D0D0; }");

    m_statusLabel = new QLabel("Ready", this);
    sb->addWidget(m_statusLabel, 1);

    m_cellSumLabel = new QLabel("", this);
    sb->addPermanentWidget(m_cellSumLabel);

    // Zoom controls
    auto* zoomFrame = new QWidget(this);
    auto* zhl = new QHBoxLayout(zoomFrame);
    zhl->setContentsMargins(4, 0, 4, 0);
    zhl->setSpacing(4);

    m_zoomLabel = new QLabel("100%", this);
    m_zoomLabel->setFixedWidth(38);

    m_zoomSlider = new QSlider(Qt::Horizontal, this);
    m_zoomSlider->setRange(50, 400);
    m_zoomSlider->setValue(100);
    m_zoomSlider->setFixedWidth(100);
    m_zoomSlider->setTickPosition(QSlider::NoTicks);

    connect(m_zoomSlider, &QSlider::valueChanged,
            this, &MainWindow::onZoomChanged);

    // Zoom in/out buttons
    auto* zoomOut = new QToolButton(this);
    zoomOut->setText("-");
    zoomOut->setFixedSize(18, 18);
    connect(zoomOut, &QToolButton::clicked, this,
            [this]{ m_zoomSlider->setValue(m_zoomSlider->value() - 10); });

    auto* zoomIn = new QToolButton(this);
    zoomIn->setText("+");
    zoomIn->setFixedSize(18, 18);
    connect(zoomIn, &QToolButton::clicked, this,
            [this]{ m_zoomSlider->setValue(m_zoomSlider->value() + 10); });

    zhl->addWidget(zoomOut);
    zhl->addWidget(m_zoomSlider);
    zhl->addWidget(zoomIn);
    zhl->addWidget(m_zoomLabel);

    sb->addPermanentWidget(zoomFrame);
}

void MainWindow::setupActions()
{
    // ── Menu bar ──────────────────────────────────────────────────────────────
    auto* mb = menuBar();

    // File
    auto* fileMenu = mb->addMenu("&File");
    fileMenu->addAction("&New",    this, &MainWindow::onNew,  QKeySequence::New);
    fileMenu->addAction("&Open...",this, &MainWindow::onOpen, QKeySequence::Open);
    fileMenu->addSeparator();
    fileMenu->addAction("&Save",   this, &MainWindow::onSave, QKeySequence::Save);
    fileMenu->addAction("Save &As...", this, &MainWindow::onSaveAs, QKeySequence::SaveAs);
    fileMenu->addSeparator();
    fileMenu->addAction("E&xit", this, &QWidget::close, QKeySequence::Quit);

    // Edit
    auto* editMenu = mb->addMenu("&Edit");
    m_undoAction = editMenu->addAction("&Undo", this, &MainWindow::onUndo, QKeySequence::Undo);
    m_redoAction = editMenu->addAction("&Redo", this, &MainWindow::onRedo, QKeySequence::Redo);
    editMenu->addSeparator();
    editMenu->addAction("Cu&t",   this, &MainWindow::onCut,   QKeySequence::Cut);
    editMenu->addAction("&Copy",  this, &MainWindow::onCopy,  QKeySequence::Copy);
    editMenu->addAction("&Paste", this, &MainWindow::onPaste, QKeySequence::Paste);
    editMenu->addSeparator();
    editMenu->addAction("&Find && Replace...", this, &MainWindow::onFindReplace, QKeySequence::Find);

    // View
    auto* viewMenu = mb->addMenu("&View");
    viewMenu->addAction("Freeze Top Row", this, [this]{
        m_view->setFreezeRow(1);
    });
    viewMenu->addAction("Freeze First Column", this, [this]{
        m_view->setFreezeCol(1);
    });
    viewMenu->addAction("Unfreeze", this, [this]{
        m_view->setFreezeRow(0); m_view->setFreezeCol(0);
    });

    // Keyboard shortcuts
    auto addShortcut = [this](const QKeySequence& seq, auto slot) {
        auto* act = new QAction(this);
        act->setShortcut(seq);
        connect(act, &QAction::triggered, this, slot);
        addAction(act);
    };

    addShortcut(QKeySequence("Ctrl+Z"), &MainWindow::onUndo);
    addShortcut(QKeySequence("Ctrl+Y"), &MainWindow::onRedo);
    addShortcut(QKeySequence("Ctrl+C"), &MainWindow::onCopy);
    addShortcut(QKeySequence("Ctrl+X"), &MainWindow::onCut);
    addShortcut(QKeySequence("Ctrl+V"), &MainWindow::onPaste);
    addShortcut(QKeySequence("Delete"), &MainWindow::onDelete);
    addShortcut(QKeySequence("Ctrl+S"), &MainWindow::onSave);
    addShortcut(QKeySequence("Ctrl+H"), &MainWindow::onFindReplace);
    addShortcut(QKeySequence("Ctrl+F"), &MainWindow::onFindReplace);
}

// ── File operations ───────────────────────────────────────────────────────────
void MainWindow::onNew()
{
    if (!maybeSave()) return;

    if (m_core) delete m_core;
    m_core = createSpreadsheetCore();
    if (!m_core) return;

    m_core->addSheet("Sheet1");
    m_view->setCore(m_core);
    m_view->setSheet(m_core->sheets().first());

    // Reset sheet tabs
    while (m_sheetBar->tabCount() > 0)
        m_sheetBar->removeTab(0);
    m_sheetBar->addTab("Sheet1");
    m_sheetBar->setCurrentTab(0);

    m_filePath.clear();
    m_modified = false;
    updateTitle();
    m_statusLabel->setText("Ready");
}

void MainWindow::onOpen()
{
    if (!maybeSave()) return;

    QString path = QFileDialog::getOpenFileName(this,
        "Open Spreadsheet", {},
        "All Supported (*.xlsx *.csv *.opensheet);;"
        "Excel Files (*.xlsx);;"
        "CSV Files (*.csv);;"
        "OpenSheet Files (*.opensheet);;"
        "All Files (*)");

    if (!path.isEmpty())
        openFile(path);
}

void MainWindow::openFile(const QString& path)
{
    m_filePath = path;
    updateTitle();
    loadFileAsync(path);
}

void MainWindow::loadFileAsync(const QString& path)
{
    // Show progress dialog
    m_progress = new QProgressDialog("Loading file...", "Cancel", 0, 100, this);
    m_progress->setWindowModality(Qt::WindowModal);
    m_progress->setMinimumDuration(500);
    m_progress->setValue(0);
    m_progress->show();

    // Create fresh core
    if (m_core) delete m_core;
    m_core = createSpreadsheetCore();
    if (!m_core) return;

    m_view->setCore(m_core);

    // Load in background thread — captures path/core by value, this by pointer
    // (safe because QFutureWatcher::finished is delivered on main thread)
    ISpreadsheetCore* core = m_core;
    auto future = QtConcurrent::run([path, core, this]() -> QString {
        IFileLoader* loader = createFileLoader();
        if (!loader) return QStringLiteral("Failed to create file loader");

        // Load metadata (sheet names, dimensions) without cell data
        loader->loadMetadata(path, core);

        auto sheets = core->sheets();
        int sheetId = sheets.isEmpty() ? 0 : sheets.first();

        // Stream-load cells — progress reported via queued signal to main thread
        bool ok = loader->loadChunk(path, core, sheetId, 0, INT_MAX,
            [this](qint64 done, qint64 total) {
                if (total > 0) {
                    int pct = (int)((double)done / total * 100.0);
                    QMetaObject::invokeMethod(this, [this, pct]{
                        if (m_progress) m_progress->setValue(pct);
                    }, Qt::QueuedConnection);
                }
            });

        QString error = ok ? QString() : loader->lastError();
        delete loader;
        return error;
    });

    // Watch for completion
    auto* watcher = new QFutureWatcher<QString>(this);
    connect(watcher, &QFutureWatcher<QString>::finished, this, [this, watcher]{
        QString err = watcher->result();
        watcher->deleteLater();
        onLoadFinished(err.isEmpty(), err);
    });
    watcher->setFuture(future);
}

void MainWindow::onLoadProgress(int percent)
{
    if (m_progress) m_progress->setValue(percent);
}

void MainWindow::onLoadFinished(bool success, const QString& error)
{
    if (m_progress) {
        m_progress->close();
        delete m_progress;
        m_progress = nullptr;
    }

    if (!success) {
        QMessageBox::critical(this, "Load Error",
            "Failed to open file:\n" + error);
        return;
    }

    // Rebuild sheet tabs
    while (m_sheetBar->tabCount() > 0)
        m_sheetBar->removeTab(0);

    if (!m_core) return;

    auto sheets = m_core->sheets();
    for (int id : sheets) {
        m_sheetBar->addTab(m_core->sheetName(id));
    }
    if (!sheets.isEmpty()) {
        m_sheetBar->setCurrentTab(0);
        m_view->setSheet(sheets.first());
    }

    m_modified = false;
    updateTitle();
    m_statusLabel->setText("File loaded successfully");
}

void MainWindow::onSave()
{
    if (m_filePath.isEmpty()) { onSaveAs(); return; }

    IFileLoader* loader = createFileLoader();
    if (!loader) {
        QMessageBox::critical(this, "Save Error", "Failed to create file saver");
        return;
    }

    auto sheets = m_core->sheets();
    int sid = sheets.isEmpty() ? 0 : sheets.first();
    bool ok = loader->save(m_filePath, m_core, sid, [](qint64, qint64){});
    delete loader;

    if (ok) {
        m_modified = false;
        updateTitle();
        m_statusLabel->setText("Saved");
    } else {
        QMessageBox::critical(this, "Save Error", "Failed to save file");
    }
}

void MainWindow::onSaveAs()
{
    QString path = QFileDialog::getSaveFileName(this,
        "Save Spreadsheet", m_filePath.isEmpty() ? "Workbook.xlsx" : m_filePath,
        "Excel Files (*.xlsx);;"
        "CSV Files (*.csv);;"
        "OpenSheet Files (*.opensheet);;"
        "All Files (*)");

    if (!path.isEmpty()) {
        m_filePath = path;
        onSave();
    }
}

// ── Edit ─────────────────────────────────────────────────────────────────────
void MainWindow::onUndo()
{
    if (m_core && m_core->canUndo()) { m_core->undo(); m_view->refresh(); }
}

void MainWindow::onRedo()
{
    if (m_core && m_core->canRedo()) { m_core->redo(); m_view->refresh(); }
}

void MainWindow::onCopy()   { if (m_view) m_view->copy(); }
void MainWindow::onCut()    { if (m_view) m_view->cut(); }
void MainWindow::onPaste()  { if (m_view) m_view->paste(); }
void MainWindow::onDelete() { if (m_view) m_view->deleteSelection(); }

// ── Cell navigation ───────────────────────────────────────────────────────────
void MainWindow::onCellChanged(int row, int col)
{
    if (!m_core) return;

    // Update name box
    QString addr = SpreadsheetView::colLabel(col) + QString::number(row + 1);
    m_formulaBar->setCellAddress(addr);

    // Update formula bar
    int sheetId = m_view->currentSheet();
    Cell cell = m_core->getCell(sheetId, row, col);
    QString formula = cell.formula.isEmpty()
                    ? (cell.rawValue.isNull() ? "" : cell.rawValue.toString())
                    : "=" + cell.formula;
    m_formulaBar->setFormulaText(formula);

    // Update status bar sum (for selection)
    auto sel = m_view->selection();
    if (!sel.isSingleCell()) {
        double sum = 0;
        int cnt = 0;
        for (int r = sel.r1; r <= sel.r2; ++r) {
            for (int c = sel.c1; c <= sel.c2; ++c) {
                Cell cc = m_core->getCell(sheetId, r, c);
                bool ok;
                double v = cc.cachedValue.toDouble(&ok);
                if (ok) { sum += v; cnt++; }
            }
        }
        if (cnt > 0)
            m_cellSumLabel->setText(QString("Sum: %1  Count: %2").arg(sum).arg(cnt));
        else
            m_cellSumLabel->clear();
    } else {
        m_cellSumLabel->clear();
    }

    // Sync ribbon
    syncRibbonToCell(row, col);
}

void MainWindow::onFormulaCommitted(const QString& formula)
{
    if (!m_core || !m_view) return;
    int row = m_view->currentRow();
    int col = m_view->currentCol();
    int sid = m_view->currentSheet();

    if (formula.startsWith('=')) {
        m_core->setCellFormula(sid, row, col, formula.mid(1));
    } else {
        bool ok;
        double d = formula.toDouble(&ok);
        if (ok) m_core->setCellValue(sid, row, col, d);
        else    m_core->setCellValue(sid, row, col, formula);
    }
    m_modified = true;
    updateTitle();
}

void MainWindow::onNameBoxNavigate(const QString& address)
{
    if (m_view) m_view->navigateToAddress(address);
}

// ── Ribbon format actions ─────────────────────────────────────────────────────
void MainWindow::syncRibbonToCell(int row, int col)
{
    if (!m_core || !m_ribbon) return;
    Cell cell = m_core->getCell(m_view->currentSheet(), row, col);
    const CellFormat& fmt = cell.format;

    RibbonFormatState state;
    state.fontFamily   = fmt.font.family().isEmpty() ? "Calibri" : fmt.font.family();
    state.fontSize     = fmt.font.pointSize() <= 0 ? 11 : fmt.font.pointSize();
    state.bold         = fmt.bold;
    state.italic       = fmt.italic;
    state.underline    = fmt.underline;
    state.textColor    = fmt.textColor;
    state.fillColor    = fmt.fillColor;
    state.wrapText     = fmt.wrapText;
    state.numberFormat = fmt.numberFormat;
    state.decimals     = fmt.decimals;

    Qt::Alignment h = fmt.alignment & Qt::AlignHorizontal_Mask;
    if (h == Qt::AlignHCenter) state.hAlign = 1;
    else if (h == Qt::AlignRight) state.hAlign = 2;
    else state.hAlign = 0;

    Qt::Alignment v = fmt.alignment & Qt::AlignVertical_Mask;
    if (v == Qt::AlignVCenter) state.vAlign = 1;
    else if (v == Qt::AlignBottom) state.vAlign = 2;
    else state.vAlign = 0;

    m_ribbon->setFormatState(state);
}

static void applyFmt(ISpreadsheetCore* core, SpreadsheetView* view,
                     std::function<void(CellFormat&)> fn)
{
    if (!core || !view) return;
    auto sel = view->selection();
    int sid = view->currentSheet();
    for (int r = sel.r1; r <= sel.r2; ++r) {
        for (int c = sel.c1; c <= sel.c2; ++c) {
            Cell cell = core->getCell(sid, r, c);
            fn(cell.format);
            core->setCellFormat(sid, r, c, cell.format);
        }
    }
    view->refresh();
}

void MainWindow::onBoldToggled(bool on)
{
    applyFmt(m_core, m_view, [on](CellFormat& f){ f.bold = on; f.font.setBold(on); });
}
void MainWindow::onItalicToggled(bool on)
{
    applyFmt(m_core, m_view, [on](CellFormat& f){ f.italic = on; f.font.setItalic(on); });
}
void MainWindow::onUnderlineToggled(bool on)
{
    applyFmt(m_core, m_view, [on](CellFormat& f){ f.underline = on; f.font.setUnderline(on); });
}
void MainWindow::onFontFamilyChanged(const QString& family)
{
    applyFmt(m_core, m_view, [&](CellFormat& f){ f.font.setFamily(family); });
}
void MainWindow::onFontSizeChanged(int sz)
{
    applyFmt(m_core, m_view, [sz](CellFormat& f){ f.font.setPointSize(sz); });
}
void MainWindow::onTextColorChanged(const QColor& c)
{
    applyFmt(m_core, m_view, [&c](CellFormat& f){ f.textColor = c; });
}
void MainWindow::onFillColorChanged(const QColor& c)
{
    applyFmt(m_core, m_view, [&c](CellFormat& f){ f.fillColor = c; });
}
void MainWindow::onHAlignChanged(int a)
{
    Qt::Alignment align;
    if (a == 1) align = Qt::AlignHCenter;
    else if (a == 2) align = Qt::AlignRight;
    else align = Qt::AlignLeft;
    applyFmt(m_core, m_view, [align](CellFormat& f){
        f.alignment = (f.alignment & Qt::AlignVertical_Mask) | align;
    });
}
void MainWindow::onVAlignChanged(int a)
{
    Qt::Alignment align;
    if (a == 1) align = Qt::AlignVCenter;
    else if (a == 2) align = Qt::AlignBottom;
    else align = Qt::AlignTop;
    applyFmt(m_core, m_view, [align](CellFormat& f){
        f.alignment = (f.alignment & Qt::AlignHorizontal_Mask) | align;
    });
}
void MainWindow::onWrapTextToggled(bool on)
{
    applyFmt(m_core, m_view, [on](CellFormat& f){ f.wrapText = on; });
}
void MainWindow::onMergeCells()
{
    if (!m_core || !m_view) return;
    auto sel = m_view->selection();
    if (sel.isSingleCell()) return;
    m_core->mergeCells(m_view->currentSheet(), sel.r1, sel.c1, sel.r2, sel.c2);
    m_view->refresh();
}
void MainWindow::onNumberFormatChanged(int fmt)
{
    applyFmt(m_core, m_view, [fmt](CellFormat& f){ f.numberFormat = fmt; });
}

// ── Row/column operations ─────────────────────────────────────────────────────
void MainWindow::onInsertRow()
{
    if (m_core && m_view)
        m_core->insertRow(m_view->currentSheet(), m_view->currentRow());
    if (m_view) m_view->refresh();
}
void MainWindow::onDeleteRow()
{
    if (m_core && m_view)
        m_core->deleteRow(m_view->currentSheet(), m_view->currentRow());
    if (m_view) m_view->refresh();
}
void MainWindow::onInsertColumn()
{
    if (m_core && m_view)
        m_core->insertColumn(m_view->currentSheet(), m_view->currentCol());
    if (m_view) m_view->refresh();
}
void MainWindow::onDeleteColumn()
{
    if (m_core && m_view)
        m_core->deleteColumn(m_view->currentSheet(), m_view->currentCol());
    if (m_view) m_view->refresh();
}

// ── AutoSum ───────────────────────────────────────────────────────────────────
void MainWindow::onAutoSum()
{
    if (!m_core || !m_view) return;
    int row = m_view->currentRow();
    int col = m_view->currentCol();
    int sid = m_view->currentSheet();

    // Look up for contiguous non-empty cells
    int topRow = row - 1;
    while (topRow >= 0) {
        Cell c = m_core->getCell(sid, topRow, col);
        if (c.rawValue.isNull()) break;
        topRow--;
    }
    topRow++;

    if (topRow < row) {
        QString formula = QString("SUM(%1%2:%1%3)")
            .arg(SpreadsheetView::colLabel(col))
            .arg(topRow + 1)
            .arg(row);
        m_core->setCellFormula(sid, row, col, formula);
        m_view->refresh();
        onCellChanged(row, col);
    }
}

// ── Sort / Filter ─────────────────────────────────────────────────────────────
void MainWindow::onSortAsc()  { if (m_view) m_view->sortAscending(); }
void MainWindow::onSortDesc() { if (m_view) m_view->sortDescending(); }
void MainWindow::onFilter()
{
    m_statusLabel->setText("Filter: select a range and use Data > Filter");
}

// ── Find & Replace ────────────────────────────────────────────────────────────
void MainWindow::onFindReplace()
{
    if (!m_findDlg) {
        m_findDlg = new FindReplaceDialog(this);
        connect(m_findDlg, &FindReplaceDialog::findNextRequested,
                this, [this](const QString& text, bool cs) {
                    if (m_view) m_view->findNext(text, cs, false);
                });
        connect(m_findDlg, &FindReplaceDialog::replaceAllRequested,
                this, [this](const QString& find, const QString& replace, bool cs) {
                    if (m_view) m_view->replaceAll(find, replace, cs, false);
                });
    }
    m_findDlg->show();
    m_findDlg->raise();
}

// ── Format cells ──────────────────────────────────────────────────────────────
void MainWindow::onFormatCells()
{
    if (!m_core || !m_view) return;
    Cell cell = m_core->getCell(m_view->currentSheet(),
                                m_view->currentRow(), m_view->currentCol());
    FormatCellsDialog dlg(cell.format, this);
    if (dlg.exec() == QDialog::Accepted) {
        applyFmt(m_core, m_view, [&dlg](CellFormat& f){ f = dlg.cellFormat(); });
    }
}

// ── Charts ────────────────────────────────────────────────────────────────────
void MainWindow::onInsertChart()
{
    m_statusLabel->setText("Chart: Select data range, then Insert > Chart");
}

// ── Sheet tabs ────────────────────────────────────────────────────────────────
void MainWindow::onSheetTabActivated(int index)
{
    if (!m_core) return;
    auto sheets = m_core->sheets();
    if (index < sheets.size()) {
        m_view->setSheet(sheets[index]);
    }
}

void MainWindow::onAddSheet()
{
    if (!m_core) return;
    QString name = QString("Sheet%1").arg(m_core->sheets().size() + 1);
    bool ok;
    name = QInputDialog::getText(this, "Insert Sheet", "Sheet name:", QLineEdit::Normal, name, &ok);
    if (!ok || name.isEmpty()) return;

    addNewSheet(name);
}

void MainWindow::addNewSheet(const QString& name)
{
    if (!m_core) return;
    int id = m_core->addSheet(name);
    int idx = m_sheetBar->addTab(name);
    m_sheetBar->setCurrentTab(idx);
    m_view->setSheet(id);
}

void MainWindow::onRenameSheet(int index)
{
    if (!m_core) return;
    auto sheets = m_core->sheets();
    if (index >= sheets.size()) return;

    bool ok;
    QString oldName = m_core->sheetName(sheets[index]);
    QString newName = QInputDialog::getText(this, "Rename Sheet",
        "New name:", QLineEdit::Normal, oldName, &ok);
    if (!ok || newName.isEmpty() || newName == oldName) return;

    m_core->renameSheet(sheets[index], newName);
    m_sheetBar->renameTab(index, newName);
}

void MainWindow::onDeleteSheet(int index)
{
    if (!m_core) return;
    auto sheets = m_core->sheets();
    if (sheets.size() <= 1) {
        QMessageBox::warning(this, "Delete Sheet",
            "Cannot delete the last sheet.");
        return;
    }
    if (index >= sheets.size()) return;

    auto reply = QMessageBox::question(this, "Delete Sheet",
        QString("Delete sheet '%1'?").arg(m_core->sheetName(sheets[index])),
        QMessageBox::Yes | QMessageBox::No);
    if (reply != QMessageBox::Yes) return;

    m_core->removeSheet(sheets[index]);
    m_sheetBar->removeTab(index);

    int newIdx = std::max(0, index - 1);
    m_sheetBar->setCurrentTab(newIdx);
    sheets = m_core->sheets();
    if (!sheets.isEmpty())
        m_view->setSheet(sheets[newIdx]);
}

// ── Zoom ─────────────────────────────────────────────────────────────────────
void MainWindow::onZoomChanged(int value)
{
    m_zoomLabel->setText(QString("%1%").arg(value));
    if (m_view) m_view->setZoom(value);
}

// ── Window ───────────────────────────────────────────────────────────────────
void MainWindow::updateTitle()
{
    QString title;
    if (!m_filePath.isEmpty()) {
        QFileInfo fi(m_filePath);
        title = fi.fileName();
    } else {
        title = "Untitled";
    }
    if (m_modified) title += " *";
    title += " — OpenSheet";
    setWindowTitle(title);
}

bool MainWindow::maybeSave()
{
    if (!m_modified) return true;
    auto reply = QMessageBox::question(this,
        "Unsaved Changes",
        "Do you want to save your changes?",
        QMessageBox::Save | QMessageBox::Discard | QMessageBox::Cancel,
        QMessageBox::Save);
    if (reply == QMessageBox::Save) { onSave(); return true; }
    if (reply == QMessageBox::Cancel) return false;
    return true;
}

void MainWindow::closeEvent(QCloseEvent* e)
{
    if (!maybeSave()) { e->ignore(); return; }

    QSettings s("OpenSheet", "OpenSheet");
    s.setValue("mainWindowGeometry", saveGeometry());
    s.setValue("mainWindowState", saveState());
    e->accept();
}

void MainWindow::dragEnterEvent(QDragEnterEvent* e)
{
    if (e->mimeData()->hasUrls()) e->acceptProposedAction();
}

void MainWindow::dropEvent(QDropEvent* e)
{
    const auto urls = e->mimeData()->urls();
    if (!urls.isEmpty()) openFile(urls.first().toLocalFile());
}
