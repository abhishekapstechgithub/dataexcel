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

class MainWindow : public QMainWindow {
    Q_OBJECT
public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

protected:
    void closeEvent(QCloseEvent* e) override;

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

private:
    ISpreadsheetCore* m_core   { nullptr };
    IFileLoader*      m_loader { nullptr };
    SpreadsheetView*  m_view   { nullptr };
    RibbonWidget*     m_ribbon { nullptr };
    QTabBar*          m_sheetBar { nullptr };
    QLabel*           m_cellRef  { nullptr };
    QLineEdit*        m_formulaBar { nullptr };
    QLabel*           m_statusLabel{ nullptr };
    QProgressBar*     m_progress   { nullptr };
    QString           m_filePath;
    bool              m_modified { false };

    void setupMenuBar();
    void setupStatusBar();
    void setupSheetBar();
    void connectRibbon();
    void rebuildSheetTabs();
    void setModified(bool m);
    bool confirmSave();
    void loadFileAsync(const QString& path);
    void updateTitle();
    SheetId currentSheet();
};
