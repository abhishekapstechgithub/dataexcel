#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  XlsxSaxLoader.h — Streaming SAX-based XLSX parser
//
//  Does NOT load the entire XML DOM into memory.
//  Instead uses Qt's QXmlStreamReader for forward-only, event-driven parsing.
//  This allows handling XLSX files of arbitrary size with bounded memory usage.
// ═══════════════════════════════════════════════════════════════════════════════
#include <QString>
#include <QStringList>
#include <QHash>
#include <functional>

class ISpreadsheetCore;
using ProgressCallback = std::function<void(qint64, qint64)>;

// ── Cell reference parser ─────────────────────────────────────────────────────
struct CellRef {
    int row { 0 };
    int col { 0 };
    static CellRef parse(const QString& ref);   // e.g. "A1" → {0,0}
};

// ── SAX XLSX loader ───────────────────────────────────────────────────────────
class XlsxSaxLoader
{
public:
    XlsxSaxLoader() = default;

    // ── Load shared strings (pass contents of sharedStrings.xml) ─────────────
    bool parseSharedStrings(const QByteArray& xml, QStringList& out);

    // ── Streaming sheet parse ─────────────────────────────────────────────────
    // Calls `onCell(row, col, value, formula)` for each non-empty cell.
    // Returns false on error. Reports progress via callback(bytesRead, total).
    bool parseSheet(const QByteArray& sheetXml,
                    const QStringList& sharedStrings,
                    const QHash<int,QString>& numFmts,
                    int firstRow, int maxRows,
                    std::function<void(int,int,const QString&,const QString&)> onCell,
                    ProgressCallback progress = {});

    // ── Load number formats (from styles.xml) ─────────────────────────────────
    bool parseStyles(const QByteArray& stylesXml, QHash<int,QString>& numFmts);

    QString lastError() const { return m_lastError; }

private:
    QString m_lastError;
};
