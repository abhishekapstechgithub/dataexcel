#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  TileCache.h — Tile-based memory cache for large spreadsheet data
//
//  Each tile covers TILE_ROWS × TILE_COLS cells.
//  Only MAX_TILES tiles are kept in RAM; oldest tiles are evicted.
//  Evicted tiles are persisted to SQLite on disk.
//
//  Thread safety: all public methods must be called from the main thread.
//  Background loading posts data via signal from worker thread.
// ═══════════════════════════════════════════════════════════════════════════════
#include <QObject>
#include <QString>
#include <QVariant>
#include <QHash>
#include <QList>
#include <QMutex>
#include <functional>

// ── Tile dimensions ───────────────────────────────────────────────────────────
static constexpr int TILE_ROWS  = 10000;   // rows per tile
static constexpr int TILE_COLS  = 500;     // cols per tile
static constexpr int MAX_TILES  = 20;      // max tiles in RAM

// ── Cell data stored in cache ─────────────────────────────────────────────────
struct CachedCell {
    QVariant value;        // display value
    QString  formula;      // raw formula (empty if plain value)
    // Formatting is kept separately in SpreadsheetEngine for active cells
};

// ── Tile key ──────────────────────────────────────────────────────────────────
struct TileKey {
    int sheet { 0 };
    int tileRow { 0 };   // row / TILE_ROWS
    int tileCol { 0 };   // col / TILE_COLS

    bool operator==(const TileKey& o) const = default;
};

inline size_t qHash(const TileKey& k, size_t seed = 0) noexcept
{
    return qHashMulti(seed, k.sheet, k.tileRow, k.tileCol);
}

// ── Tile data ─────────────────────────────────────────────────────────────────
struct Tile {
    TileKey                             key;
    QHash<quint32, CachedCell>          cells;   // row*TILE_COLS+col → cell
    qint64                              lruStamp { 0 };
    bool                                dirty    { false };

    CachedCell* get(int localRow, int localCol) {
        quint32 k = (quint32)localRow * TILE_COLS + localCol;
        auto it = cells.find(k);
        return it == cells.end() ? nullptr : &(*it);
    }

    void set(int localRow, int localCol, const CachedCell& c) {
        quint32 k = (quint32)localRow * TILE_COLS + localCol;
        cells[k] = c;
        dirty = true;
    }
};

// ═════════════════════════════════════════════════════════════════════════════
class TileCache : public QObject
{
    Q_OBJECT
public:
    explicit TileCache(QObject* parent = nullptr);
    ~TileCache() override;

    // ── Initialise with a backing SQLite file (pass ":memory:" for RAM-only) ─
    bool open(const QString& dbPath);
    void close();

    // ── Cell access ──────────────────────────────────────────────────────────
    // Returns nullptr if tile not loaded (triggers async load)
    const CachedCell* getCell(int sheet, int row, int col);

    // Synchronous write (marks tile dirty)
    void setCell(int sheet, int row, int col, const CachedCell& cell);

    // Bulk-write from background loader (called from any thread via queued conn)
    // chunks: list of (sheet, row, col, value, formula)
    using BulkChunk = std::tuple<int,int,int,QVariant,QString>;
    void bulkInsert(const QList<BulkChunk>& chunk);

    // ── Cache management ─────────────────────────────────────────────────────
    void flush();             // write all dirty tiles to SQLite
    void evictAll();          // flush + clear RAM
    void invalidate(int sheet); // discard all tiles for a sheet

    // ── Dimensions tracking ──────────────────────────────────────────────────
    qint64 maxRow(int sheet) const;
    int    maxCol(int sheet) const;

signals:
    // Emitted when a tile has been loaded from disk (so view can repaint)
    void tileLoaded(int sheet, int tileRow, int tileCol);

private:
    Tile*  getOrLoadTile(int sheet, int tileRow, int tileCol);
    void   evictOldest();
    void   saveTile(Tile* tile);
    void   loadTile(Tile* tile, int sheet, int tileRow, int tileCol);

    bool   ensureTable(int sheet);

    QHash<TileKey, Tile*>  m_tiles;
    QList<TileKey>         m_lruOrder;
    qint64                 m_stamp { 0 };

    struct Impl;
    Impl*  d { nullptr };

    QHash<int, qint64> m_maxRow;
    QHash<int, int>    m_maxCol;
};
