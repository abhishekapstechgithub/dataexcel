#include "IFileLoader.h"
#include "CsvLoader.h"
#include "XlsxLoader.h"
#include <QFileInfo>

class FileLoaderImpl : public IFileLoader {
    Q_OBJECT
public:
    bool loadChunk(const QString& filePath, ISpreadsheetCore* target,
                   SheetId sheet, int firstRow, int count,
                   ProgressCallback progress) override
    {
        QString ext = QFileInfo(filePath).suffix().toLower();
        if (ext == "csv") {
            CsvLoader csv;
            bool ok = csv.loadChunk(filePath, target, sheet, firstRow, count, progress);
            m_error = csv.lastError();
            emit loadFinished(ok);
            return ok;
        } else if (ext == "xlsx") {
            XlsxLoader xlsx;
            bool ok = xlsx.load(filePath, target, firstRow, count, progress);
            m_error = xlsx.lastError();
            emit loadFinished(ok);
            return ok;
        }
        m_error = "Unsupported format: " + ext;
        emit loadFinished(false);
        return false;
    }

    bool loadMetadata(const QString& filePath, ISpreadsheetCore* target) override {
        // For CSV: count lines quickly; for XLSX: parse workbook.xml
        QString ext = QFileInfo(filePath).suffix().toLower();
        if (ext == "csv") {
            QFile f(filePath);
            if (!f.open(QIODevice::ReadOnly)) { m_error = f.errorString(); return false; }
            int rows = 0;
            while (!f.atEnd()) { f.readLine(); rows++; }
            SheetId s = target->sheets().isEmpty() ? target->addSheet() : target->sheets().first();
            Q_UNUSED(s);
        }
        return true;
    }

    bool save(const QString& filePath, ISpreadsheetCore* source,
              SheetId sheet, ProgressCallback progress) override
    {
        QString ext = QFileInfo(filePath).suffix().toLower();
        if (ext == "csv") {
            CsvLoader csv;
            bool ok = csv.save(filePath, source, sheet, progress);
            m_error = csv.lastError();
            emit saveFinished(ok);
            return ok;
        } else if (ext == "xlsx") {
            XlsxLoader xlsx;
            bool ok = xlsx.save(filePath, source, sheet, progress);
            m_error = xlsx.lastError();
            emit saveFinished(ok);
            return ok;
        }
        m_error = "Unsupported format: " + ext;
        emit saveFinished(false);
        return false;
    }

    QStringList supportedReadFormats()  const override { return {"csv","xlsx"}; }
    QStringList supportedWriteFormats() const override { return {"csv","xlsx"}; }
    QString lastError() const override { return m_error; }

private:
    QString m_error;
};

#include "FileLoader.moc"

extern "C" LOADER_API IFileLoader* createFileLoader() {
    return new FileLoaderImpl();
}
