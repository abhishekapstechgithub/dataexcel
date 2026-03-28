#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  MainWindow.h — Main application window (Excel-like layout)
// ═══════════════════════════════════════════════════════════════════════════════
#include <QMainWindow>
#include <QString>
#include <QColor>
#include <memory>

class ISpreadsheetCore;
class RibbonWidget;
class FormulaBar;
class SpreadsheetView;
class SheetTabBar;
class FindReplaceDialog;
class FormatCellsDialog;
class QLabel;
class QSlider;
class QToolButton;
class QProgressDialog;
class PluginManager;

class MainWindow : public QMainWindow
{
    Q_OBJECT
public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

    // Open a file (called from main() with command-line arg)
    void openFile(const QString& path);

protected:
    void closeEvent(QCloseEvent* e) override;
    void dragEnterEvent(QDragEnterEvent* e) override;
    void dropEvent(QDropEvent* e) override;

private slots:
    // ── File ─────────────────────────────────────────────────────────────────
    void onNew();
    void onOpen();
    void onSave();
    void onSaveAs();

    // ── Edit ──────────────────────────────────────────────────────────────────
    void onUndo();
    void onRedo();
    void onCopy();
    void onCut();
    void onPaste();
    void onDelete();

    // ── Cell navigation ───────────────────────────────────────────────────────
    void onCellChanged(int row, int col);
    void onFormulaCommitted(const QString& formula);
    void onNameBoxNavigate(const QString& address);

    // ── Ribbon signals ────────────────────────────────────────────────────────
    void onBoldToggled(bool on);
    void onItalicToggled(bool on);
    void onUnderlineToggled(bool on);
    void onFontFamilyChanged(const QString& family);
    void onFontSizeChanged(int size);
    void onTextColorChanged(const QColor& c);
    void onFillColorChanged(const QColor& c);
    void onHAlignChanged(int a);
    void onVAlignChanged(int a);
    void onWrapTextToggled(bool on);
    void onMergeCells();
    void onNumberFormatChanged(int fmt);
    void onInsertRow();
    void onDeleteRow();
    void onInsertColumn();
    void onDeleteColumn();
    void onAutoSum();
    void onSortAsc();
    void onSortDesc();
    void onFilter();
    void onFindReplace();
    void onFormatCells();
    void onInsertChart();

    // ── Sheet tabs ────────────────────────────────────────────────────────────
    void onSheetTabActivated(int index);
    void onAddSheet();
    void onRenameSheet(int index);
    void onDeleteSheet(int index);

    // ── Zoom / status ─────────────────────────────────────────────────────────
    void onZoomChanged(int value);

    // ── File loading ──────────────────────────────────────────────────────────
    void onLoadProgress(int percent);
    void onLoadFinished(bool success, const QString& error);

private:
    void setupRibbon();
    void setupFormulaBar();
    void setupView();
    void setupSheetTabs();
    void setupStatusBar();
    void setupActions();
    void syncRibbonToCell(int row, int col);
    void updateTitle();
    bool maybeSave();
    void loadFileAsync(const QString& path);
    void addNewSheet(const QString& name = {});
    QString currentFilePath() const { return m_filePath; }

    // ── Widgets ───────────────────────────────────────────────────────────────
    RibbonWidget*      m_ribbon    { nullptr };
    FormulaBar*        m_formulaBar { nullptr };
    SpreadsheetView*   m_view      { nullptr };
    SheetTabBar*       m_sheetBar  { nullptr };
    FindReplaceDialog* m_findDlg   { nullptr };

    // Status bar widgets
    QLabel*      m_statusLabel  { nullptr };
    QLabel*      m_cellSumLabel { nullptr };
    QLabel*      m_zoomLabel    { nullptr };
    QSlider*     m_zoomSlider   { nullptr };

    // ── Data ─────────────────────────────────────────────────────────────────
    ISpreadsheetCore*  m_core     { nullptr };
    QString            m_filePath;
    bool               m_modified { false };

    // ── Async loading ─────────────────────────────────────────────────────────
    QProgressDialog* m_progress { nullptr };

    // ── Plugin system ─────────────────────────────────────────────────────────
    PluginManager*   m_pluginManager { nullptr };

    // ── Actions ───────────────────────────────────────────────────────────────
    QAction* m_undoAction { nullptr };
    QAction* m_redoAction { nullptr };
};
