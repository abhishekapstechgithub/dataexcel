#pragma once
#ifdef _WIN32
  #ifdef FILELOADER_EXPORTS
    #define LOADER_API __declspec(dllexport)
  #else
    #define LOADER_API __declspec(dllimport)
  #endif
#else
  #define LOADER_API __attribute__((visibility("default")))
#endif

#include <QObject>
#include <QString>
#include "ISpreadsheetCore.h"

// Progress callback: (bytesRead, totalBytes) — called from worker thread
using ProgressCallback = std::function<void(qint64, qint64)>;

class LOADER_API IFileLoader : public QObject
{
    Q_OBJECT
public:
    virtual ~IFileLoader() = default;

    // ── Streaming load ───────────────────────────────────────────────────────
    // Loads ONLY rows [firstRow, firstRow+count) of the given sheet.
    // Use firstRow=0, count=INT_MAX to load the whole file (not recommended for large files).
    virtual bool loadChunk(const QString& filePath,
                           ISpreadsheetCore* target,
                           SheetId sheet,
                           int firstRow,
                           int count,
                           ProgressCallback progress = {}) = 0;

    // Load metadata only (sheet names, row/col counts) without loading cell data
    virtual bool loadMetadata(const QString& filePath,
                              ISpreadsheetCore* target) = 0;

    // ── Save ─────────────────────────────────────────────────────────────────
    virtual bool save(const QString& filePath,
                      ISpreadsheetCore* source,
                      SheetId sheet,
                      ProgressCallback progress = {}) = 0;

    // ── Supported formats ────────────────────────────────────────────────────
    virtual QStringList supportedReadFormats()  const = 0;
    virtual QStringList supportedWriteFormats() const = 0;

    virtual QString lastError() const = 0;

signals:
    void loadProgress(qint64 bytesRead, qint64 total);
    void loadFinished(bool success);
    void saveFinished(bool success);
};

extern "C" LOADER_API IFileLoader* createFileLoader();
