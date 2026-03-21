#pragma once
#include "ISpreadsheetCore.h"
#include <QString>
#include <functional>

class CsvLoader {
public:
    bool loadChunk(const QString& path, ISpreadsheetCore* core, SheetId sheet,
                   int firstRow, int count, std::function<void(qint64,qint64)> progress);
    bool save(const QString& path, ISpreadsheetCore* core, SheetId sheet,
              std::function<void(qint64,qint64)> progress);
    QString lastError() const { return m_error; }
private:
    QString m_error;
};
