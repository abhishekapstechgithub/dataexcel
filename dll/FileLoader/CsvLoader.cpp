#include "CsvLoader.h"
#include <QFile>
#include <QTextStream>
#include <QStringList>
#include <QStringConverter>

// Parse one CSV line respecting quoted fields
static QStringList parseCsvLine(const QString& line) {
    QStringList fields;
    QString field;
    bool inQuote = false;
    for (int i = 0; i < line.size(); ++i) {
        QChar c = line[i];
        if (c == '"') {
            if (inQuote && i+1 < line.size() && line[i+1] == '"') {
                field += '"'; i++;
            } else inQuote = !inQuote;
        } else if (c == ',' && !inQuote) {
            fields << field; field.clear();
        } else field += c;
    }
    fields << field;
    return fields;
}

bool CsvLoader::loadChunk(const QString& path, ISpreadsheetCore* core,
                          SheetId sheet, int firstRow, int count,
                          std::function<void(qint64,qint64)> progress)
{
    QFile f(path);
    if (!f.open(QIODevice::ReadOnly | QIODevice::Text)) {
        m_error = f.errorString(); return false;
    }
    qint64 total = f.size();
    QTextStream in(&f);
    in.setEncoding(QStringConverter::Utf8);

    int rowIndex = 0, loaded = 0;
    while (!in.atEnd() && loaded < count) {
        QString line = in.readLine();
        if (rowIndex < firstRow) { rowIndex++; continue; }
        QStringList fields = parseCsvLine(line);
        for (int c = 0; c < fields.size(); ++c)
            core->setCellValue(sheet, rowIndex, c, fields[c]);
        rowIndex++; loaded++;
        if (progress && (loaded % 1000 == 0))
            progress(f.pos(), total);
    }
    if (progress) progress(total, total);
    return true;
}

bool CsvLoader::save(const QString& path, ISpreadsheetCore* core,
                     SheetId sheet, std::function<void(qint64,qint64)> progress)
{
    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Text)) {
        m_error = f.errorString(); return false;
    }
    QTextStream out(&f);
    int rows = core->rowCount(sheet);
    int cols = core->columnCount(sheet);
    for (int r = 0; r < rows; ++r) {
        QStringList line;
        bool hasData = false;
        for (int c = 0; c < cols; ++c) {
            Cell cell = core->getCell(sheet, r, c);
            QString v = cell.cachedValue.isValid() ? cell.cachedValue.toString()
                                                   : cell.rawValue.toString();
            if (!v.isEmpty()) hasData = true;
            if (v.contains(',') || v.contains('"') || v.contains('\n'))
                v = '"' + v.replace("\"","\"\"") + '"';
            line << v;
        }
        if (!hasData) continue;
        // trim trailing empty
        while (!line.isEmpty() && line.last().isEmpty()) line.removeLast();
        out << line.join(',') << '\n';
        if (progress && r % 10000 == 0) progress(r, rows);
    }
    if (progress) progress(rows, rows);
    return true;
}
