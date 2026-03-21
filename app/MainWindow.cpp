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
#include <QApplication>

MainWindow::MainWindow(QWidget* parent) : QMainWindow(parent) {
    setWindowTitle("OpenSheet");
    resize(1280, 800);

    // Load DLLs via factory functions
    m_core   = createSpreadsheetCore();
    m_loader = createFileLoader();

    // ── Central widget ───────────────────────────────────────────────────────
    auto* central = new QWidget;
    auto* vl      = new QVBoxLayout(central);
    vl->setContentsMargins(0,0,0,0);
    vl->setSpacing(0);

    // Ribbon
    m_ribbon = createRibbonWidget(this);
    vl->addWidget(m_ribbon);

    // Formula bar
    auto* fbarLayout = new QHBoxLayout;
    fbarLayout->setContentsMargins(4,2,4,2);
    m_cellRef = new QLabel("A1");
    m_cellRef->setFixedWidth(60);
    m_cellRef->setStyleSheet("border:1px solid #ccc; padding:2px; font-size:12px;");
    m_formulaBar = new QLineEdit;
    m_formulaBar->setPlaceholderText("Enter value or formula...");
    m_formulaBar->setStyleSheet("font-size:12px; padding:2px;");
    fbarLayout->addWidget(m_cellRef);
    fbarLayout->addWidget(m_formulaBar);
    vl->addLayout(fbarLayout);

    // Spreadsheet view
    m_view = new SpreadsheetView(m_core, m_core->sheets().first(), this);
    vl->addWidget(m_view, 1);

    // Sheet tab bar
    setupSheetBar();
    vl->addWidget(m_sheetBar);

    setCentralWidget(central);

    setupMenuBar();
    setupStatusBar();
    connectRibbon();
    rebuildSheetTabs();
    updateTitle();

    // Formula bar → cell edit
    connect(m_formulaBar, &QLineEdit::returnPressed, this, [this]() {
        QString text = m_formulaBar->text();
        int r = m_view->currentRow(), c = m_view->currentCol();
        if (text.startsWith('='))
            m_core->setCellFormula(currentSheet(), r, c, text);
        else
            m_core->setCellValue(currentSheet(), r, c, text);
        setModified(true);
    });

    // Selection → formula bar
    connect(m_view, &SpreadsheetView::selectionFormatChanged,
            this, &MainWindow::onFormatChanged);
}

MainWindow::~MainWindow() {
    delete m_core;
    delete m_loader;
}

SheetId MainWindow::currentSheet() {
    auto sheets = m_core->sheets();
    int idx = m_sheetBar->currentIndex();
    return (idx >= 0 && idx < sheets.size()) ? sheets[idx] : sheets.first();
}

void MainWindow::setupMenuBar() {
    auto* mb = menuBar();
    mb->setStyleSheet("QMenuBar { background:#217346; color:white; font-size:12px; }"
                      "QMenuBar::item:selected { background:#185c37; }"
                      "QMenu { background:white; color:black; }");

    // File
    auto* fileMenu = mb->addMenu("&File");
    fileMenu->addAction("&New",     this, &MainWindow::newFile,  QKeySequence::New);
    fileMenu->addAction("&Open...", this, &MainWindow::openFile, QKeySequence::Open);
    fileMenu->addSeparator();
    fileMenu->addAction("&Save",        this, &MainWindow::saveFile,   QKeySequence::Save);
    fileMenu->addAction("Save &As...",  this, &MainWindow::saveFileAs, QKeySequence::SaveAs);
    fileMenu->addSeparator();
    fileMenu->addAction("E&xit", qApp, &QApplication::quit, QKeySequence::Quit);

    // Edit
    auto* editMenu = mb->addMenu("&Edit");
    editMenu->addAction("&Undo", m_core, &ISpreadsheetCore::undo, QKeySequence::Undo);
    editMenu->addAction("&Redo", m_core, &ISpreadsheetCore::redo, QKeySequence::Redo);
    editMenu->addSeparator();
    editMenu->addAction("Find && &Replace...", this, [this]{
        auto* dlg = new FindReplaceDialog(m_core, currentSheet(), this);
        dlg->exec();
    }, QKeySequence(Qt::CTRL | Qt::Key_H));

    // Sheet
    auto* sheetMenu = mb->addMenu("&Sheet");
    sheetMenu->addAction("Add Sheet",    this, &MainWindow::addSheet);
    sheetMenu->addAction("Remove Sheet", this, &MainWindow::removeSheet);
}

void MainWindow::setupStatusBar() {
    m_statusLabel = new QLabel("Ready");
    m_progress    = new QProgressBar;
    m_progress->setFixedWidth(180);
    m_progress->setVisible(false);
    statusBar()->addWidget(m_statusLabel, 1);
    statusBar()->addPermanentWidget(m_progress);
    statusBar()->setStyleSheet("QStatusBar { font-size:11px; }");
}

void MainWindow::setupSheetBar() {
    m_sheetBar = new QTabBar;
    m_sheetBar->setTabsClosable(false);
    m_sheetBar->setMovable(true);
    m_sheetBar->setStyleSheet(
        "QTabBar::tab { min-width:70px; padding:4px 10px; font-size:11px;"
        "  border:1px solid #ccc; border-bottom:none; background:#e8e8e8; }"
        "QTabBar::tab:selected { background:white; font-weight:bold; }"
    );
    connect(m_sheetBar, &QTabBar::currentChanged, this, &MainWindow::switchSheet);
    connect(m_sheetBar, &QTabBar::tabBarDoubleClicked, this, &MainWindow::renameSheet);
}

void MainWindow::rebuildSheetTabs() {
    QSignalBlocker b(m_sheetBar);
    while (m_sheetBar->count() > 0) m_sheetBar->removeTab(0);
    for (SheetId s : m_core->sheets())
        m_sheetBar->addTab(m_core->sheetName(s));
}

void MainWindow::connectRibbon() {
    // File
    connect(m_ribbon, &RibbonWidget::newFileRequested,  this, &MainWindow::newFile);
    connect(m_ribbon, &RibbonWidget::openFileRequested, this, &MainWindow::openFile);
    connect(m_ribbon, &RibbonWidget::saveFileRequested, this, &MainWindow::saveFile);
    connect(m_ribbon, &RibbonWidget::saveAsRequested,   this, &MainWindow::saveFileAs);
    // Clipboard
    connect(m_ribbon, &RibbonWidget::pasteRequested, m_view, [this]{
        QApplication::sendEvent(m_view, new QKeyEvent(QEvent::KeyPress, Qt::Key_V, Qt::ControlModifier));
    });
    connect(m_ribbon, &RibbonWidget::cutRequested, this, [this]{
        m_core->clearCell(currentSheet(), m_view->currentRow(), m_view->currentCol());
    });
    connect(m_ribbon, &RibbonWidget::copyRequested, m_view, [this]{
        QApplication::sendEvent(m_view, new QKeyEvent(QEvent::KeyPress, Qt::Key_C, Qt::ControlModifier));
    });
    // Font
    connect(m_ribbon, &RibbonWidget::fontFamilyChanged, m_view, &SpreadsheetView::applyFontFamily);
    connect(m_ribbon, &RibbonWidget::fontSizeChanged,   m_view, &SpreadsheetView::applyFontSize);
    connect(m_ribbon, &RibbonWidget::boldToggled,       m_view, &SpreadsheetView::applyBold);
    connect(m_ribbon, &RibbonWidget::italicToggled,     m_view, &SpreadsheetView::applyItalic);
    connect(m_ribbon, &RibbonWidget::underlineToggled,  m_view, &SpreadsheetView::applyUnderline);
    connect(m_ribbon, &RibbonWidget::textColorChanged,  m_view, &SpreadsheetView::applyTextColor);
    connect(m_ribbon, &RibbonWidget::fillColorChanged,  m_view, &SpreadsheetView::applyFillColor);
    // Alignment
    connect(m_ribbon, &RibbonWidget::hAlignChanged,       m_view, &SpreadsheetView::applyHAlign);
    connect(m_ribbon, &RibbonWidget::vAlignChanged,       m_view, &SpreadsheetView::applyVAlign);
    connect(m_ribbon, &RibbonWidget::wrapTextToggled,     m_view, &SpreadsheetView::applyWrapText);
    connect(m_ribbon, &RibbonWidget::mergeCellsRequested, m_view, &SpreadsheetView::mergeSelected);
    // Number
    connect(m_ribbon, &RibbonWidget::numberFormatChanged, m_view, &SpreadsheetView::applyNumberFormat);
    connect(m_ribbon, &RibbonWidget::increaseDecimalRequested, this, [this]{
        auto fmt = m_view->currentCellFormat(); fmt.decimals++; m_view->applyFormat(fmt);
    });
    connect(m_ribbon, &RibbonWidget::decreaseDecimalRequested, this, [this]{
        auto fmt = m_view->currentCellFormat();
        fmt.decimals = qMax(0, fmt.decimals-1); m_view->applyFormat(fmt);
    });
    // Cells
    connect(m_ribbon, &RibbonWidget::insertRowRequested,    m_view, &SpreadsheetView::insertRow);
    connect(m_ribbon, &RibbonWidget::deleteRowRequested,    m_view, &SpreadsheetView::deleteRow);
    connect(m_ribbon, &RibbonWidget::insertColumnRequested, m_view, &SpreadsheetView::insertColumn);
    connect(m_ribbon, &RibbonWidget::deleteColumnRequested, m_view, &SpreadsheetView::deleteColumn);
    connect(m_ribbon, &RibbonWidget::formatCellsRequested, this, [this]{
        FormatCellsDialog dlg(m_view->currentCellFormat(), this);
        if (dlg.exec() == QDialog::Accepted) m_view->applyFormat(dlg.result());
    });
    // Editing
    connect(m_ribbon, &RibbonWidget::autoSumRequested,    m_view, &SpreadsheetView::autoSum);
    connect(m_ribbon, &RibbonWidget::sortAscRequested,    m_view, &SpreadsheetView::sortAsc);
    connect(m_ribbon, &RibbonWidget::sortDescRequested,   m_view, &SpreadsheetView::sortDesc);
    connect(m_ribbon, &RibbonWidget::findReplaceRequested, this, [this]{
        auto* dlg = new FindReplaceDialog(m_core, currentSheet(), this);
        dlg->setAttribute(Qt::WA_DeleteOnClose);
        dlg->show();
    });
}

void MainWindow::onFormatChanged(const CellFormat& fmt, const QString& ref) {
    m_cellRef->setText(ref);
    Cell cell = m_core->getCell(currentSheet(),
                                m_view->currentRow(), m_view->currentCol());
    m_formulaBar->setText(cell.formula.isEmpty()
        ? cell.rawValue.toString() : cell.formula);

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
    m_ribbon->setFormatState(state);
}

void MainWindow::newFile() {
    if (!confirmSave()) return;
    delete m_core;
    m_core = createSpreadsheetCore();
    m_filePath.clear();
    setModified(false);
    m_view->switchSheet(m_core->sheets().first());
    rebuildSheetTabs();
    updateTitle();
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
    m_statusLabel->setText("Loading " + QFileInfo(path).fileName() + "...");

    auto unused_future = QtConcurrent::run([this, path]() {
        m_loader->loadChunk(path, m_core, m_core->sheets().first(),
                            0, INT_MAX,
                            [this](qint64 read, qint64 total) {
            if (total > 0)
                QMetaObject::invokeMethod(m_progress, "setValue", Qt::QueuedConnection,
                    Q_ARG(int, (int)(read * 100 / total)));
        });
        QMetaObject::invokeMethod(this, [this, path]() {
            m_progress->setVisible(false);
            m_statusLabel->setText("Loaded: " + QFileInfo(path).fileName());
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
    m_statusLabel->setText("Saved: " + QFileInfo(m_filePath).fileName());
}

void MainWindow::saveFileAs() {
    QString path = QFileDialog::getSaveFileName(this, "Save As", m_filePath,
        "CSV Files (*.csv);;Excel Files (*.xlsx)");
    if (path.isEmpty()) return;
    m_filePath = path;
    saveFile();
}

void MainWindow::addSheet() {
    bool ok;
    QString name = QInputDialog::getText(this, "New Sheet", "Sheet name:",
                                         QLineEdit::Normal, {}, &ok);
    if (!ok) return;
    SheetId s = m_core->addSheet(name);
    m_sheetBar->addTab(m_core->sheetName(s));
    m_sheetBar->setCurrentIndex(m_sheetBar->count()-1);
}

void MainWindow::removeSheet() {
    auto sheets = m_core->sheets();
    if (sheets.size() <= 1) {
        QMessageBox::warning(this, "Remove Sheet", "Cannot remove the last sheet.");
        return;
    }
    int idx = m_sheetBar->currentIndex();
    if (idx < 0 || idx >= sheets.size()) return;
    if (QMessageBox::question(this, "Remove Sheet",
            QString("Remove \"%1\"?").arg(m_core->sheetName(sheets[idx])),
            QMessageBox::Yes|QMessageBox::No) != QMessageBox::Yes) return;
    m_core->removeSheet(sheets[idx]);
    rebuildSheetTabs();
    m_view->switchSheet(currentSheet());
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
    }
}

void MainWindow::switchSheet(int idx) {
    auto sheets = m_core->sheets();
    if (idx >= 0 && idx < sheets.size())
        m_view->switchSheet(sheets[idx]);
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
    if (btn == QMessageBox::Save) { saveFile(); return true; }
    if (btn == QMessageBox::Discard) return true;
    return false;
}

void MainWindow::updateTitle() {
    QString name = m_filePath.isEmpty() ? "Untitled" : QFileInfo(m_filePath).fileName();
    setWindowTitle((m_modified ? "* " : "") + name + " — OpenSheet");
}

void MainWindow::closeEvent(QCloseEvent* e) {
    if (confirmSave()) e->accept();
    else e->ignore();
}
