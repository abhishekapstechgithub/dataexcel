#pragma once
#include <QDialog>
#include <QString>

class QLineEdit;
class QCheckBox;

// ── FindReplaceDialog ─────────────────────────────────────────────────────────
// Emits signals; the main window connects them to SpreadsheetView::findNext /
// replaceAll so this dialog has no direct dependency on the core.
class FindReplaceDialog : public QDialog
{
    Q_OBJECT
public:
    explicit FindReplaceDialog(QWidget* parent = nullptr);

signals:
    void findNextRequested   (const QString& text, bool caseSensitive);
    void replaceAllRequested (const QString& find, const QString& replace, bool caseSensitive);

private slots:
    void onFindNext();
    void onReplaceAll();

private:
    QLineEdit* m_findEdit    { nullptr };
    QLineEdit* m_replaceEdit { nullptr };
    QCheckBox* m_matchCase   { nullptr };
};
