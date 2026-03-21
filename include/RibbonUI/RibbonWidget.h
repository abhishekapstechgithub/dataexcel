#pragma once
#ifdef _WIN32
  #ifdef RIBBONUI_EXPORTS
    #define RIBBON_API __declspec(dllexport)
  #else
    #define RIBBON_API __declspec(dllimport)
  #endif
#else
  #define RIBBON_API __attribute__((visibility("default")))
#endif

#include <QWidget>
#include <QFont>
#include <QColor>

// Forward
class ISpreadsheetCore;

// ── Format state for the ribbon to reflect / apply ───────────────────────────
struct RibbonFormatState {
    QString fontFamily   { "Calibri" };
    int     fontSize     { 11 };
    bool    bold         { false };
    bool    italic       { false };
    bool    underline    { false };
    QColor  textColor    { Qt::black };
    QColor  fillColor    { Qt::white };
    int     hAlign       { 0 };  // 0=Left 1=Center 2=Right
    int     vAlign       { 0 };  // 0=Top  1=Mid    2=Bottom
    bool    wrapText     { false };
    int     numberFormat { 0 };  // 0=General 1=Number 2=Currency 3=Percent 4=Scientific
    int     decimals     { 2 };
};

class RIBBON_API RibbonWidget : public QWidget
{
    Q_OBJECT
public:
    explicit RibbonWidget(QWidget* parent = nullptr);
    ~RibbonWidget() override;

    // Sync ribbon state to reflect the currently selected cell
    void setFormatState(const RibbonFormatState& state);

signals:
    // ── Clipboard ────────────────────────────────────────────────────────────
    void pasteRequested();
    void cutRequested();
    void copyRequested();
    void formatPainterRequested();

    // ── Font ─────────────────────────────────────────────────────────────────
    void fontFamilyChanged(const QString& family);
    void fontSizeChanged(int size);
    void boldToggled(bool on);
    void italicToggled(bool on);
    void underlineToggled(bool on);
    void textColorChanged(const QColor& color);
    void fillColorChanged(const QColor& color);

    // ── Alignment ────────────────────────────────────────────────────────────
    void hAlignChanged(int align);   // 0=L 1=C 2=R
    void vAlignChanged(int align);   // 0=T 1=M 2=B
    void wrapTextToggled(bool on);
    void mergeCellsRequested();

    // ── Number format ────────────────────────────────────────────────────────
    void numberFormatChanged(int fmt);
    void increaseDecimalRequested();
    void decreaseDecimalRequested();

    // ── Cells ────────────────────────────────────────────────────────────────
    void insertRowRequested();
    void deleteRowRequested();
    void insertColumnRequested();
    void deleteColumnRequested();
    void formatCellsRequested();

    // ── Editing ──────────────────────────────────────────────────────────────
    void autoSumRequested();
    void sortAscRequested();
    void sortDescRequested();
    void filterRequested();
    void findReplaceRequested();

    // ── File ─────────────────────────────────────────────────────────────────
    void newFileRequested();
    void openFileRequested();
    void saveFileRequested();
    void saveAsRequested();

private:
    struct Impl;
    Impl* d;
};

extern "C" RIBBON_API RibbonWidget* createRibbonWidget(QWidget* parent = nullptr);
