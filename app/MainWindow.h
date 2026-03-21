#pragma once
#include <QMainWindow>
#include "ISpreadsheetCore.h"
#include "IFileLoader.h"
#include "RibbonWidget.h"

class SpreadsheetView;
class QTabBar;
class QLabel;
class QLineEdit;
class QProgressBar;
class QSlider;
class QToolButton;
class QSplitter;
class QStackedWidget;

// ── MainWindow ────────────────────────────────────────────────────────────────
// Excel-style application window:
//   [Title bar with Quick Access Toolbar]
//   [Ribbon - tabbed groups: Home/Insert/Page Layout/Formulas/Data/Review/View]
//   [Formula bar: [Name Box] [fx] [Input]]
//   [Sheet area: QTableView with row/col headers]
//   [Sheet tab bar + status bar + zoom slider]
class MainWindow : public QMainWindow {
    Q_OBJECT
public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

protected:
    void closeEvent(QCloseEvent* e) override;
    void resizeEvent(QResizeEvent* e) override;

private slots:
    void newFile();
    void openFile();
    void saveFile();
    void saveFileAs();
    void onFormatChanged(const CellFormat& fmt, const QString& ref);
    void addSheet();
    void removeSheet();
    void renameSheet(int idx);
    void switchSheet(int idx);
    void onZoomChanged(int value);
    void onFormulaBarReturn();
    void onSelectionChanged(const CellFormat& fmt, const QString& ref);
    void updateSelectionStats();

private:
    // ── Core components ───────────────────────────────────────────────────────
    ISpreadsheetCore* m_core   { nullptr };
    IFileLoader*      m_loader { nullptr };
    SpreadsheetView*  m_view   { nullptr };
    RibbonWidget*     m_ribbon { nullptr };

    // ── Formula bar widgets ───────────────────────────────────────────────────
    QLineEdit*   m_nameBox    { nullptr };  // "A1" cell reference
    QLabel*      m_fxLabel    { nullptr };  // "fx" icon
    QLineEdit*   m_formulaBar { nullptr };  // formula/value input

    // ── Sheet tab bar ─────────────────────────────────────────────────────────
    QTabBar*    m_sheetBar   { nullptr };
    QToolButton* m_addSheetBtn { nullptr };

    // ── Status bar widgets ────────────────────────────────────────────────────
    QLabel*      m_statusReady  { nullptr };
    QLabel*      m_statusMode   { nullptr };
    QLabel*      m_statusStats  { nullptr };
    QProgressBar* m_progress    { nullptr };

    // ── Zoom controls ─────────────────────────────────────────────────────────
    QSlider*     m_zoomSlider   { nullptr };
    QLabel*      m_zoomLabel    { nullptr };
    int          m_zoomPercent  { 100 };

    // ── State ─────────────────────────────────────────────────────────────────
    QString      m_filePath;
    bool         m_modified     { false };
    bool         m_formulaEditing { false };

    // ── Setup helpers ─────────────────────────────────────────────────────────
    void setupQuickAccessBar();
    void setupFormulaBar();
    void setupSheetBar();
    void setupStatusBar();
    void connectRibbon();
    void connectView();
    void rebuildSheetTabs();
    void setModified(bool m);
    bool confirmSave();
    void loadFileAsync(const QString& path);
    void updateTitle();
    void setZoom(int percent);
    SheetId currentSheet() const;

    // ── Quick Access Toolbar items ────────────────────────────────────────────
    QToolButton* makeQATButton(const QString& text, const QString& tip,
                               const std::function<void()>& fn);
};
