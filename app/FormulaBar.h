#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  FormulaBar.h — Excel-style formula bar
//  Shows cell address, fx button, and formula/value editor
// ═══════════════════════════════════════════════════════════════════════════════
#include <QWidget>
#include <QString>

class QLineEdit;
class QToolButton;
class QLabel;
class QComboBox;

class FormulaBar : public QWidget
{
    Q_OBJECT
public:
    explicit FormulaBar(QWidget* parent = nullptr);
    ~FormulaBar() override = default;

    // Set the cell address label (e.g. "A1", "B12")
    void setCellAddress(const QString& address);

    // Set the formula/value text shown in the edit field
    void setFormulaText(const QString& text);

    // Get the current formula text
    QString formulaText() const;

    // Clear formula bar
    void clear();

    // Set editing mode (formula bar gets keyboard focus)
    void beginEdit();
    void endEdit();

signals:
    // Emitted when user presses Enter or navigates away
    void formulaCommitted(const QString& formula);

    // Emitted when user presses Escape
    void editCancelled();

    // Emitted when the name box address is entered (e.g. user types "B5" + Enter)
    void nameBoxNavigate(const QString& address);

    // Emitted when fx button clicked
    void insertFunctionRequested();

protected:
    bool eventFilter(QObject* obj, QEvent* event) override;

private slots:
    void onNameBoxReturnPressed();
    void onFormulaEditReturnPressed();
    void onFormulaEditEscapePressed();

private:
    QLineEdit*   m_nameBox   { nullptr };
    QToolButton* m_fxButton  { nullptr };
    QLineEdit*   m_formulaEdit { nullptr };
    bool         m_editing   { false };
};
