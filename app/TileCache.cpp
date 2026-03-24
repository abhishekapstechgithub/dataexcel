// ═══════════════════════════════════════════════════════════════════════════════
//  TileCache.cpp — Tile-based cache for 100GB file support
// ═══════════════════════════════════════════════════════════════════════════════
#include "TileCache.h"
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QUuid>
#include <algorithm>

struct TileCache::Impl {
    QSqlDatabase db;
    QString      connectionName;
    bool         open { false };
};

TileCache::TileCache(QObject* parent)
    : QObject(parent)
    , d(new Impl)
{
    d->connectionName = QStringLiteral("TileCache_") + QUuid::createUuid().toString(QUuid::Id128);
}

TileCache::~TileCache()
{
    close();
    qDeleteAll(m_tiles);
    delete d;
}

bool TileCache::open(const QString& dbPath)
{
    d->db = QSqlDatabase::addDatabase("QSQLITE", d->connectionName);
    d->db.setDatabaseName(dbPath);

    if (!d->db.open()) {
        qWarning() << "TileCache: cannot open SQLite:" << d->db.lastError().text();
        return false;
    }

    // Enable WAL for performance
    QSqlQuery q(d->db);
    q.exec("PRAGMA journal_mode=WAL");
    q.exec("PRAGMA synchronous=NORMAL");
    q.exec("PRAGMA cache_size=8000");

    d->open = true;
    return true;
}

void TileCache::close()
{
    if (!d->open) return;
    flush();
    d->db.close();
    QSqlDatabase::removeDatabase(d->connectionName);
    d->open = false;
}

bool TileCache::ensureTable(int sheet)
{
    if (!d->open) return false;
    QSqlQuery q(d->db);
    QString tbl = QString("sheet_%1").arg(sheet);
    q.exec(QString(
        "CREATE TABLE IF NOT EXISTS %1 ("
        "  trow INTEGER,"
        "  tcol INTEGER,"
        "  lrow INTEGER,"
        "  lcol INTEGER,"
        "  value TEXT,"
        "  formula TEXT,"
        "  PRIMARY KEY(trow,tcol,lrow,lcol)"
        ")").arg(tbl));
    if (q.lastError().isValid()) {
        qWarning() << "TileCache ensureTable:" << q.lastError().text();
        return false;
    }
    return true;
}

void TileCache::loadTile(Tile* tile, int sheet, int tileRow, int tileCol)
{
    if (!d->open) return;
    ensureTable(sheet);

    QSqlQuery q(d->db);
    q.prepare(QString("SELECT lrow,lcol,value,formula FROM sheet_%1 WHERE trow=? AND tcol=?").arg(sheet));
    q.addBindValue(tileRow);
    q.addBindValue(tileCol);
    q.exec();

    while (q.next()) {
        int lrow = q.value(0).toInt();
        int lcol = q.value(1).toInt();
        CachedCell cc;
        cc.value   = q.value(2);
        cc.formula = q.value(3).toString();
        tile->cells[(quint32)lrow * TILE_COLS + lcol] = cc;
    }
    tile->dirty = false;
}

void TileCache::saveTile(Tile* tile)
{
    if (!d->open || !tile->dirty) return;
    ensureTable(tile->key.sheet);

    QSqlQuery q(d->db);
    // Delete old data for this tile
    q.prepare(QString("DELETE FROM sheet_%1 WHERE trow=? AND tcol=?").arg(tile->key.sheet));
    q.addBindValue(tile->key.tileRow);
    q.addBindValue(tile->key.tileCol);
    q.exec();

    // Insert new
    d->db.transaction();
    q.prepare(QString("INSERT INTO sheet_%1 (trow,tcol,lrow,lcol,value,formula) VALUES(?,?,?,?,?,?)").arg(tile->key.sheet));

    for (auto it = tile->cells.constBegin(); it != tile->cells.constEnd(); ++it) {
        quint32 key = it.key();
        int lrow = (int)(key / TILE_COLS);
        int lcol = (int)(key % TILE_COLS);
        q.addBindValue(tile->key.tileRow);
        q.addBindValue(tile->key.tileCol);
        q.addBindValue(lrow);
        q.addBindValue(lcol);
        q.addBindValue(it.value().value.toString());
        q.addBindValue(it.value().formula);
        q.exec();
    }
    d->db.commit();
    tile->dirty = false;
}

void TileCache::evictOldest()
{
    if (m_tiles.isEmpty()) return;

    // Find tile with lowest LRU stamp that is not the current
    TileKey oldest;
    qint64 minStamp = std::numeric_limits<qint64>::max();
    for (auto it = m_tiles.constBegin(); it != m_tiles.constEnd(); ++it) {
        if (it.value()->lruStamp < minStamp) {
            minStamp = it.value()->lruStamp;
            oldest = it.key();
        }
    }

    Tile* t = m_tiles.take(oldest);
    saveTile(t);
    m_lruOrder.removeOne(oldest);
    delete t;
}

Tile* TileCache::getOrLoadTile(int sheet, int tileRow, int tileCol)
{
    TileKey key { sheet, tileRow, tileCol };
    auto it = m_tiles.find(key);
    if (it != m_tiles.end()) {
        it.value()->lruStamp = ++m_stamp;
        return it.value();
    }

    // Evict if at capacity
    if (m_tiles.size() >= MAX_TILES)
        evictOldest();

    Tile* tile = new Tile;
    tile->key      = key;
    tile->lruStamp = ++m_stamp;
    loadTile(tile, sheet, tileRow, tileCol);
    m_tiles[key]   = tile;
    m_lruOrder.append(key);
    return tile;
}

const CachedCell* TileCache::getCell(int sheet, int row, int col)
{
    int tileRow  = row / TILE_ROWS;
    int tileCol  = col / TILE_COLS;
    int localRow = row % TILE_ROWS;
    int localCol = col % TILE_COLS;

    Tile* tile = getOrLoadTile(sheet, tileRow, tileCol);
    return tile->get(localRow, localCol);
}

void TileCache::setCell(int sheet, int row, int col, const CachedCell& cell)
{
    int tileRow  = row / TILE_ROWS;
    int tileCol  = col / TILE_COLS;
    int localRow = row % TILE_ROWS;
    int localCol = col % TILE_COLS;

    Tile* tile = getOrLoadTile(sheet, tileRow, tileCol);
    tile->set(localRow, localCol, cell);

    // Update dimensions
    if (!m_maxRow.contains(sheet) || row > m_maxRow[sheet])
        m_maxRow[sheet] = row;
    if (!m_maxCol.contains(sheet) || col > m_maxCol[sheet])
        m_maxCol[sheet] = col;
}

void TileCache::bulkInsert(const QList<BulkChunk>& chunks)
{
    // Group by tile
    struct TileWrite {
        QList<std::tuple<int,int,CachedCell>> cells;
    };
    QHash<TileKey, TileWrite> groups;

    for (const auto& [sheet, row, col, value, formula] : chunks) {
        int tileRow  = row / TILE_ROWS;
        int tileCol  = col / TILE_COLS;
        int localRow = row % TILE_ROWS;
        int localCol = col % TILE_COLS;
        CachedCell cc { value, formula };

        groups[TileKey{sheet, tileRow, tileCol}].cells.emplace_back(localRow, localCol, cc);

        if (!m_maxRow.contains(sheet) || row > m_maxRow[sheet])
            m_maxRow[sheet] = row;
        if (!m_maxCol.contains(sheet) || col > m_maxCol[sheet])
            m_maxCol[sheet] = col;
    }

    for (auto it = groups.begin(); it != groups.end(); ++it) {
        const TileKey& k = it.key();
        Tile* tile = getOrLoadTile(k.sheet, k.tileRow, k.tileCol);
        for (const auto& [lr, lc, cc] : it.value().cells)
            tile->set(lr, lc, cc);
    }
}

void TileCache::flush()
{
    for (auto* tile : m_tiles) {
        saveTile(tile);
    }
}

void TileCache::evictAll()
{
    flush();
    qDeleteAll(m_tiles);
    m_tiles.clear();
    m_lruOrder.clear();
}

void TileCache::invalidate(int sheet)
{
    QList<TileKey> toRemove;
    for (auto it = m_tiles.constBegin(); it != m_tiles.constEnd(); ++it) {
        if (it.key().sheet == sheet)
            toRemove << it.key();
    }
    for (const auto& k : toRemove) {
        delete m_tiles.take(k);
        m_lruOrder.removeOne(k);
    }
    m_maxRow.remove(sheet);
    m_maxCol.remove(sheet);
}

qint64 TileCache::maxRow(int sheet) const
{
    return m_maxRow.value(sheet, 0);
}

int TileCache::maxCol(int sheet) const
{
    return m_maxCol.value(sheet, 0);
}
