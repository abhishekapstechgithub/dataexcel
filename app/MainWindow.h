#pragma once
#include <QMainWindow>
#include <functional>
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

// ── MainWindow ────────────────────────────────────────────────────────────────
// OpenSheet application shell — architecture mirrors WPS Office:
//   et.exe (shell) → etcore.dll (core) → griddrawer.dll (grid) →
//   xlsxrw.dll (io) → ksoui.dll (ribbon)
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
    void onSelectionChanged(const CellFormat& fmt, const QString& ref);
    void onFormulaBarReturn();
    void onZoomChanged(int value);
    void updateSelectionStats();
    void addSheet();
    void removeSheet();
    void renameSheet(int idx);
    void switchSheet(int idx);

private:
    // Core components
    ISpreadsheetCore* m_core   { nullptr };
    IFileLoader*      m_loader { nullptr };
    SpreadsheetView*  m_view   { nullptr };
    RibbonWidget*     m_ribbon { nullptr };

    // Formula bar
    QLineEdit* m_nameBox    { nullptr };
    QLabel*    m_fxLabel    { nullptr };
    QLineEdit* m_formulaBar { nullptr };

    // Sheet tab bar
    QTabBar*     m_sheetBar    { nullptr };
    QToolButton* m_addSheetBtn { nullptr };

    // Status bar
    QLabel*       m_statusReady { nullptr };
    QLabel*       m_statusMode  { nullptr };
    QLabel*       m_statusStats { nullptr };
    QProgressBar* m_progress    { nullptr };
    QWidget*      m_notifBar   { nullptr };

    // Zoom
    QSlider* m_zoomSlider  { nullptr };
    QLabel*  m_zoomLabel   { nullptr };
    int      m_zoomPercent { 100 };

    // State
    QString m_filePath;
    bool    m_modified       { false };
    bool    m_formulaEditing { false };

    // Builders (return widget for layout insertion)
    void     buildMenuBar();
    QWidget* buildFormulaBar();
    QWidget* buildNotifBar();
    QWidget* buildSheetBar();
    void     buildStatusBar();

    void connectRibbon();
    void connectView();
    void rebuildSheetTabs();
    void setModified(bool m);
    bool confirmSave();
    void loadFileAsync(const QString& path);
    void updateTitle();
    void setZoom(int percent);
    SheetId currentSheet() const;

    QToolButton* makeQATButton(const QString& text, const QString& tip,
                               const std::function<void()>& fn);
};
