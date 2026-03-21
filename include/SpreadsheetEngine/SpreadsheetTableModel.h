#pragma once
#ifdef _WIN32
  #ifdef SPREADSHEETENGINE_EXPORTS
    #define ENGINE_API __declspec(dllexport)
  #else
    #define ENGINE_API __declspec(dllimport)
  #endif
#else
  #define ENGINE_API __attribute__((visibility("default")))
#endif

#include <QAbstractTableModel>
#include <QObject>
#include "ISpreadsheetCore.h"

// ── Virtual table model ───────────────────────────────────────────────────────
// QTableView talks to this model.  Only visible cells are ever queried —
// the model never loads the whole sheet into Qt item storage.
class ENGINE_API SpreadsheetTableModel : public QAbstractTableModel
{
    Q_OBJECT
public:
    explicit SpreadsheetTableModel(ISpreadsheetCore* core,
                                   SheetId sheet,
                                   QObject* parent = nullptr);
    ~SpreadsheetTableModel() override;

    // ── QAbstractTableModel overrides ────────────────────────────────────────
    int     rowCount   (const QModelIndex& parent = {}) const override;
    int     columnCount(const QModelIndex& parent = {}) const override;
    QVariant data      (const QModelIndex& index, int role = Qt::DisplayRole) const override;
    bool    setData    (const QModelIndex& index, const QVariant& value,
                        int role = Qt::EditRole) override;
    QVariant headerData(int section, Qt::Orientation orientation,
                        int role = Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;

    // ── Sheet switch ─────────────────────────────────────────────────────────
    void setSheet(SheetId sheet);
    SheetId currentSheet() const;

    // ── Utility ──────────────────────────────────────────────────────────────
    static QString columnLabel(int col);   // 0→A, 25→Z, 26→AA …
    static int     columnIndex(const QString& label); // "A"→0

private slots:
    void onCellChanged(SheetId sheet, int row, int col);

private:
    struct Impl;
    Impl* d;
};

// Factory
extern "C" ENGINE_API SpreadsheetTableModel*
    createSpreadsheetModel(ISpreadsheetCore* core, SheetId sheet);
