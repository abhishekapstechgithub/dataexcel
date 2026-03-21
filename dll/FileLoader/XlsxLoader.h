#pragma once
#include "ISpreadsheetCore.h"
#include <QString>
#include <functional>

// Minimal XLSX reader/writer using Qt XML (no external libs required)
class XlsxLoader {
public:
    bool load(const QString& path, ISpreadsheetCore* core,
              int firstRow, int count,
              std::function<void(qint64,qint64)> progress);
    bool save(const QString& path, ISpreadsheetCore* core, SheetId sheet,
              std::function<void(qint64,qint64)> progress);
    QString lastError() const { return m_error; }
private:
    QString m_error;
    QStringList m_sharedStrings;
    bool parseSharedStrings(const QByteArray& xml);
    bool parseSheet(const QByteArray& xml, ISpreadsheetCore* core,
                    SheetId sheet, int firstRow, int count,
                    std::function<void(qint64,qint64)> progress);
    static QPair<int,int> cellRefToRC(const QString& ref);
};
