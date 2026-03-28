// ─────────────────────────────────────────────────────────────────────────────
//  PythonEngine.cpp — Qt-side IPC client for the Python/PySpark data engine.
//
//  Starts the Python IPC server as a child process (QProcess), then connects
//  to it over a TCP loopback socket.  Each public method serialises a JSON
//  request with an incrementing ID and emits resultReady() when the matching
//  response arrives from a background reader thread.
// ─────────────────────────────────────────────────────────────────────────────
#include "IPythonEngine.h"
#include <QProcess>
#include <QTcpSocket>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QThread>
#include <QMutex>
#include <QMutexLocker>
#include <QTimer>
#include <QCoreApplication>
#include <QDir>
#include <QFile>
#include <atomic>

// ── Reader thread ─────────────────────────────────────────────────────────────
// Reads newline-delimited JSON responses from the socket and emits signals.
class SocketReader : public QThread {
    Q_OBJECT
public:
    explicit SocketReader(QTcpSocket* sock, QObject* parent = nullptr)
        : QThread(parent), m_sock(sock) {}

    void stop() { m_stop = true; }

signals:
    void responseReceived(const QByteArray& line);

protected:
    void run() override {
        QByteArray buf;
        while (!m_stop) {
            if (!m_sock->waitForReadyRead(200)) continue;
            buf += m_sock->readAll();
            while (buf.contains('\n')) {
                int idx = buf.indexOf('\n');
                QByteArray line = buf.left(idx).trimmed();
                buf = buf.mid(idx + 1);
                if (!line.isEmpty())
                    emit responseReceived(line);
            }
        }
    }

private:
    QTcpSocket* m_sock;
    std::atomic<bool> m_stop { false };
};

// ── PythonEngine implementation ───────────────────────────────────────────────
class PythonEngineImpl : public IPythonEngine {
    Q_OBJECT
public:
    explicit PythonEngineImpl(QObject* parent = nullptr)
        : IPythonEngine(parent) {}

    ~PythonEngineImpl() override { stop(); }

    // ── Life-cycle ────────────────────────────────────────────────────────────
    bool start(const QString& pythonExe, int port) override {
        m_port = port;

        // Locate the IPC server script
        QString script = QCoreApplication::applicationDirPath()
                       + "/../python/data_engine/ipc_server.py";
        if (!QFile::exists(script)) {
            // Try beside the executable
            script = QCoreApplication::applicationDirPath()
                   + "/data_engine/ipc_server.py";
        }

        m_process = new QProcess(this);
        m_process->setProgram(pythonExe);
        m_process->setArguments({ "-m", "data_engine.ipc_server",
                                   "--port", QString::number(port) });

        // Set working directory to the python/ folder
        QString pyDir = QCoreApplication::applicationDirPath() + "/../python";
        if (QDir(pyDir).exists()) m_process->setWorkingDirectory(pyDir);

        connect(m_process, &QProcess::errorOccurred,
                this, [this](QProcess::ProcessError) {
            emit engineError("Python process error: " + m_process->errorString());
        });

        m_process->start();
        if (!m_process->waitForStarted(5000)) {
            m_error = "Failed to start Python process";
            return false;
        }

        // Give the server a moment to bind the socket
        QThread::msleep(800);
        return connectSocket();
    }

    void stop() override {
        if (m_reader) { m_reader->stop(); m_reader->wait(2000); delete m_reader; m_reader = nullptr; }
        if (m_sock)   { m_sock->disconnectFromHost(); delete m_sock; m_sock = nullptr; }
        if (m_process){ m_process->terminate(); m_process->waitForFinished(3000); delete m_process; m_process = nullptr; }
        m_running = false;
    }

    bool isRunning() const override { return m_running; }

    // ── Requests ──────────────────────────────────────────────────────────────
    int ping() override { return send({{"action","ping"}}); }

    int loadCsv(const QString& path, int chunk, int chunkSize) override {
        return send({{"action","load_csv"},{"path",path},
                     {"chunk",chunk},{"chunk_size",chunkSize}});
    }
    int loadParquet(const QString& path, int chunk, int chunkSize) override {
        return send({{"action","load_parquet"},{"path",path},
                     {"chunk",chunk},{"chunk_size",chunkSize}});
    }
    int getChunk(int dfId, int chunk, int chunkSize) override {
        return send({{"action","get_chunk"},{"df_id",dfId},
                     {"chunk",chunk},{"chunk_size",chunkSize}});
    }
    int aggregate(int dfId, const QVariantList& operations) override {
        QJsonObject req;
        req["action"]     = "aggregate";
        req["df_id"]      = dfId;
        req["operations"] = QJsonArray::fromVariantList(operations);
        return sendJson(req);
    }
    int filterRows(int dfId, const QString& col, const QString& op,
                   const QVariant& value, int chunk, int chunkSize) override {
        return send({{"action","filter"},{"df_id",dfId},{"column",col},
                     {"operator",op},{"value",value.toString()},
                     {"chunk",chunk},{"chunk_size",chunkSize}});
    }
    int releaseDataFrame(int dfId) override {
        return send({{"action","release"},{"df_id",dfId}});
    }

    QString lastError() const override { return m_error; }

private:
    bool connectSocket() {
        m_sock = new QTcpSocket(this);
        m_sock->connectToHost("127.0.0.1", m_port);
        if (!m_sock->waitForConnected(3000)) {
            m_error = "Cannot connect to Python IPC server: " + m_sock->errorString();
            delete m_sock; m_sock = nullptr;
            return false;
        }
        m_reader = new SocketReader(m_sock, this);
        connect(m_reader, &SocketReader::responseReceived,
                this,     &PythonEngineImpl::onResponse,
                Qt::QueuedConnection);
        m_reader->start();
        m_running = true;
        return true;
    }

    int send(const QVariantMap& params) {
        QJsonObject obj = QJsonObject::fromVariantMap(params);
        return sendJson(obj);
    }

    int sendJson(QJsonObject obj) {
        if (!m_sock || !m_running) {
            m_error = "Engine not running";
            return -1;
        }
        int id = ++m_nextId;
        obj["id"] = id;
        QByteArray data = QJsonDocument(obj).toJson(QJsonDocument::Compact) + "\n";
        QMutexLocker lock(&m_writeMutex);
        m_sock->write(data);
        m_sock->flush();
        return id;
    }

private slots:
    void onResponse(const QByteArray& line) {
        QJsonParseError err;
        QJsonDocument doc = QJsonDocument::fromJson(line, &err);
        if (doc.isNull()) return;
        QVariantMap result = doc.object().toVariantMap();
        int id = result.value("id").toInt();
        emit resultReady(id, result);
    }

private:
    QProcess*     m_process { nullptr };
    QTcpSocket*   m_sock    { nullptr };
    SocketReader* m_reader  { nullptr };
    QMutex        m_writeMutex;
    int           m_port    { 9876 };
    bool          m_running { false };
    QString       m_error;
    std::atomic<int> m_nextId { 0 };
};

#include "PythonEngine.moc"

extern "C" PYENG_API IPythonEngine* createPythonEngine() {
    return new PythonEngineImpl();
}
