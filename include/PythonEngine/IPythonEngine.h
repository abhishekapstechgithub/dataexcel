#pragma once
// ─────────────────────────────────────────────────────────────────────────────
//  IPythonEngine — interface to the Python/PySpark big-data backend.
//
//  The Qt app starts the Python IPC server as a child process, then
//  communicates with it over a TCP loopback socket using line-delimited JSON.
//
//  All methods are asynchronous: they return a request ID immediately and emit
//  resultReady(id, result) when the response arrives on a background thread.
// ─────────────────────────────────────────────────────────────────────────────
#ifdef _WIN32
  #ifdef PYTHONENGINE_EXPORTS
    #define PYENG_API __declspec(dllexport)
  #else
    #define PYENG_API __declspec(dllimport)
  #endif
#else
  #define PYENG_API __attribute__((visibility("default")))
#endif

#include <QObject>
#include <QString>
#include <QVariantMap>

class PYENG_API IPythonEngine : public QObject
{
    Q_OBJECT
public:
    virtual ~IPythonEngine() = default;

    // ── Life-cycle ─────────────────────────────────────────────────────────────
    // Start the Python IPC server subprocess and connect to it.
    // pythonExe: path to python / python3 executable.
    // port:      TCP port (default 9876).
    virtual bool start(const QString& pythonExe = "python3", int port = 9876) = 0;
    virtual void stop() = 0;
    virtual bool isRunning() const = 0;

    // ── Async requests ─────────────────────────────────────────────────────────
    // Each method returns a request ID.  resultReady(id, result) is emitted
    // on the calling thread (via queued connection) when the response is ready.

    virtual int ping() = 0;

    // Load a CSV file.  Returns a df_id in the result map.
    virtual int loadCsv(const QString& path, int chunk = 0, int chunkSize = 1000) = 0;

    // Load a Parquet file.
    virtual int loadParquet(const QString& path, int chunk = 0, int chunkSize = 1000) = 0;

    // Fetch a subsequent chunk from an already-loaded df.
    virtual int getChunk(int dfId, int chunk, int chunkSize = 1000) = 0;

    // Aggregations: operations = list of {"type":"sum","column":"Price"} maps.
    virtual int aggregate(int dfId, const QVariantList& operations) = 0;

    // Filter rows.  operator: "=", "!=", ">", "<", ">=", "<="
    virtual int filterRows(int dfId, const QString& column,
                           const QString& op, const QVariant& value,
                           int chunk = 0, int chunkSize = 1000) = 0;

    // Release a cached DataFrame on the Python side.
    virtual int releaseDataFrame(int dfId) = 0;

    virtual QString lastError() const = 0;

signals:
    // Emitted when any async response arrives.
    void resultReady(int requestId, const QVariantMap& result);
    // Emitted if the subprocess crashes or the socket disconnects.
    void engineError(const QString& message);
};

extern "C" PYENG_API IPythonEngine* createPythonEngine();
