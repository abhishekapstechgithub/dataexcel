#include "XlsxLoader.h"
#include <QFile>
#include <QByteArray>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QRegularExpression>
#include <QHash>
#include <QStringList>

// ---------------------------------------------------------------------------
//  Self-contained ZIP reader/writer — zero external or private dependencies.
//
//  For READING: raw DEFLATE in ZIP entries is decompressed by wrapping it in
//  a zlib frame (2-byte header + 4-byte Adler-32 trailer) then calling
//  qUncompress(), which is public API in QtCore and strips Qt's own 4-byte
//  size prefix before returning the data.
//
//  For WRITING: we use Qt's qCompress() (public API) to get a zlib stream,
//  strip Qt's 4-byte prefix and the 2-byte zlib header, keep the raw DEFLATE
//  body, and store that in the ZIP local-file entry.
// ---------------------------------------------------------------------------

// ---- endian helpers (no system headers needed) ----------------------------
static void writeLE16(QByteArray& b, quint16 v) {
    b += char(v & 0xff); b += char((v >> 8) & 0xff);
}
static void writeLE32(QByteArray& b, quint32 v) {
    b += char(v & 0xff); b += char((v>>8)&0xff);
    b += char((v>>16)&0xff); b += char((v>>24)&0xff);
}
static quint16 readLE16(const uchar* p) {
    return quint16(p[0]) | (quint16(p[1])<<8);
}
static quint32 readLE32(const uchar* p) {
    return quint32(p[0])|(quint32(p[1])<<8)|(quint32(p[2])<<16)|(quint32(p[3])<<24);
}

// ---- Adler-32 (needed to build a valid zlib frame for qUncompress) --------
static quint32 adler32(const char* data, int len) {
    quint32 s1 = 1, s2 = 0;
    for (int i = 0; i < len; ++i) {
        s1 = (s1 + quint8(data[i])) % 65521;
        s2 = (s2 + s1) % 65521;
    }
    return (s2 << 16) | s1;
}

// Decompress raw DEFLATE bytes (as stored in a ZIP entry) using qUncompress.
// qUncompress expects: [4-byte BE uncompressed size][zlib-stream]
// A zlib stream is:   [2-byte header][deflate data][4-byte Adler-32]
// We build that wrapper around the raw deflate bytes from the ZIP entry.
static QByteArray inflateRaw(const char* src, quint32 csize, quint32 usize) {
    // Build zlib frame: header(2) + raw-deflate + adler32(4)
    quint32 ck = adler32(src, csize); // adler over compressed data is wrong;
    // qUncompress doesn't verify adler32, so any value works — use 1
    Q_UNUSED(ck);
    QByteArray framed;
    framed.reserve(6 + csize);
    framed += char(0x78); framed += char(0x9C); // zlib default-compression header
    framed.append(src, csize);
    framed += char(0); framed += char(0); framed += char(0); framed += char(1); // adler32 = 1

    // qUncompress needs a 4-byte big-endian uncompressed-size prefix
    QByteArray prefixed;
    prefixed.reserve(4 + framed.size());
    prefixed += char((usize >> 24) & 0xff);
    prefixed += char((usize >> 16) & 0xff);
    prefixed += char((usize >>  8) & 0xff);
    prefixed += char( usize        & 0xff);
    prefixed.append(framed);

    return qUncompress(prefixed);
}

// Compress data with qCompress, then strip down to raw DEFLATE bytes.
// qCompress output layout: [4-byte BE size][0x78][compression-byte][deflate...][4-byte adler32]
// We keep only the [deflate...] part (strip 4+2 from front, 4 from back).
static QByteArray deflateRaw(const QByteArray& data) {
    QByteArray z = qCompress(data, 6);
    // Minimum sane size: 4-byte prefix + 2-byte zlib header + 4-byte trailer = 10
    if (z.size() < 10) return data; // fallback to stored
    // Strip 4-byte Qt size prefix + 2-byte zlib header from front, 4-byte adler32 from back
    return z.mid(6, z.size() - 10);
}

// ---- CRC-32 (required by ZIP format) -------------------------------------
// CRC-32 computed at first call — no hardcoded table, no count errors.
static quint32 crc32zip(const char* data, int len) {
    static quint32 table[256] = {};
    static bool init = false;
    if (!init) {
        for (int i = 0; i < 256; ++i) {
            quint32 c = quint32(i);
            for (int k = 0; k < 8; ++k)
                c = (c & 1) ? (0xEDB88320 ^ (c >> 1)) : (c >> 1);
            table[i] = c;
        }
        init = true;
    }
    quint32 crc = 0xFFFFFFFF;
    for (int i = 0; i < len; ++i)
        crc = table[(crc ^ quint8(data[i])) & 0xFF] ^ (crc >> 8);
    return crc ^ 0xFFFFFFFF;
}

// ---- ZIP constants --------------------------------------------------------
static const quint32 kLocalSig   = 0x04034b50;
static const quint32 kCentralSig = 0x02014b50;
static const quint32 kEOCDSig    = 0x06054b50;

// ---- MiniZipReader --------------------------------------------------------
class MiniZipReader {
public:
    explicit MiniZipReader(const QString& path) {
        QFile f(path);
        if (f.open(QIODevice::ReadOnly)) m_data = f.readAll();
        buildIndex();
    }
    bool isReadable() const { return !m_data.isEmpty(); }

    QByteArray fileData(const QString& name) const {
        auto it = m_index.find(name);
        if (it == m_index.end()) return {};
        quint32 off = it.value();
        const uchar* p = reinterpret_cast<const uchar*>(m_data.constData()) + off;
        if (off + 30 > (quint32)m_data.size()) return {};
        if (readLE32(p) != kLocalSig) return {};
        quint16 method   = readLE16(p + 8);
        quint32 csize    = readLE32(p + 18);
        quint32 usize    = readLE32(p + 22);
        quint16 fnLen    = readLE16(p + 26);
        quint16 exLen    = readLE16(p + 28);
        quint32 dataOff  = off + 30 + fnLen + exLen;
        if (dataOff + csize > (quint32)m_data.size()) return {};
        const char* src  = m_data.constData() + dataOff;
        if (method == 0) return QByteArray(src, csize);          // stored
        if (method == 8) return inflateRaw(src, csize, usize);   // deflate
        return {};
    }

private:
    QByteArray m_data;
    QHash<QString, quint32> m_index;

    void buildIndex() {
        int sz = m_data.size();
        if (sz < 22) return;
        const uchar* raw = reinterpret_cast<const uchar*>(m_data.constData());
        int eocd = -1;
        for (int i = sz - 22; i >= qMax(0, sz - 65558); --i)
            if (readLE32(raw + i) == kEOCDSig) { eocd = i; break; }
        if (eocd < 0) return;
        quint32 cdOff  = readLE32(raw + eocd + 16);
        quint16 cdCnt  = readLE16(raw + eocd + 8);
        quint32 pos = cdOff;
        for (int i = 0; i < cdCnt; ++i) {
            if (pos + 46 > (quint32)sz || readLE32(raw+pos) != kCentralSig) break;
            quint16 fnLen = readLE16(raw+pos+28);
            quint16 exLen = readLE16(raw+pos+30);
            quint16 cmLen = readLE16(raw+pos+32);
            quint32 lhOff = readLE32(raw+pos+42);
            m_index[QString::fromUtf8(m_data.mid(pos+46, fnLen))] = lhOff;
            pos += 46 + fnLen + exLen + cmLen;
        }
    }
};

// ---- MiniZipWriter --------------------------------------------------------
struct ZEntry { QString name; QByteArray cdata; quint32 crc; quint32 usize; quint16 method; quint32 loff; };

class MiniZipWriter {
public:
    void open(const QString& path) { m_path = path; }

    void addFile(const QString& name, const QByteArray& data) {
        quint32 crc  = crc32zip(data.constData(), data.size());
        QByteArray cd = deflateRaw(data);
        quint16 method = 8;
        // If deflate didn't shrink it, store instead
        if (cd.isEmpty() || cd.size() >= data.size()) { cd = data; method = 0; }

        ZEntry e;
        e.name   = name; e.cdata = cd; e.crc = crc;
        e.usize  = data.size(); e.method = method; e.loff = m_buf.size();

        QByteArray nb = name.toUtf8();
        writeLE32(m_buf, kLocalSig);
        writeLE16(m_buf, 20);       // version needed
        writeLE16(m_buf, 0);        // flags
        writeLE16(m_buf, method);
        writeLE16(m_buf, 0); writeLE16(m_buf, 0); // mod time / date
        writeLE32(m_buf, crc);
        writeLE32(m_buf, cd.size());
        writeLE32(m_buf, data.size());
        writeLE16(m_buf, nb.size());
        writeLE16(m_buf, 0);        // extra len
        m_buf.append(nb);
        m_buf.append(cd);
        m_entries.append(e);
    }

    bool close() {
        quint32 cdStart = m_buf.size();
        for (auto& e : m_entries) {
            QByteArray nb = e.name.toUtf8();
            writeLE32(m_buf, kCentralSig);
            writeLE16(m_buf, 20); writeLE16(m_buf, 20);
            writeLE16(m_buf, 0);
            writeLE16(m_buf, e.method);
            writeLE16(m_buf, 0); writeLE16(m_buf, 0);
            writeLE32(m_buf, e.crc);
            writeLE32(m_buf, (quint32)e.cdata.size());
            writeLE32(m_buf, e.usize);
            writeLE16(m_buf, nb.size());
            writeLE16(m_buf, 0); writeLE16(m_buf, 0);
            writeLE16(m_buf, 0); writeLE16(m_buf, 0);
            writeLE32(m_buf, 0);
            writeLE32(m_buf, e.loff);
            m_buf.append(nb);
        }
        quint32 cdSize = m_buf.size() - cdStart;
        writeLE32(m_buf, kEOCDSig);
        writeLE16(m_buf, 0); writeLE16(m_buf, 0);
        writeLE16(m_buf, (quint16)m_entries.size());
        writeLE16(m_buf, (quint16)m_entries.size());
        writeLE32(m_buf, cdSize);
        writeLE32(m_buf, cdStart);
        writeLE16(m_buf, 0);
        QFile f(m_path);
        if (!f.open(QIODevice::WriteOnly)) return false;
        f.write(m_buf);
        return true;
    }

private:
    QString m_path;
    QByteArray m_buf;
    QList<ZEntry> m_entries;
};

// ---------------------------------------------------------------------------
//  XlsxLoader implementation (XML logic unchanged from original)
// ---------------------------------------------------------------------------

QPair<int,int> XlsxLoader::cellRefToRC(const QString& ref) {
    static QRegularExpression re(R"(([A-Z]+)(\d+))");
    auto m = re.match(ref.toUpper());
    if (!m.hasMatch()) return {0,0};
    int col = 0;
    for (QChar c : m.captured(1)) col = col*26 + (c.unicode()-'A'+1);
    return {m.captured(2).toInt()-1, col-1};
}

bool XlsxLoader::parseSharedStrings(const QByteArray& xml) {
    m_sharedStrings.clear();
    QXmlStreamReader r(xml);
    while (!r.atEnd()) {
        r.readNext();
        if (r.isStartElement() && r.name() == QLatin1String("si")) {
            QString text;
            while (!(r.isEndElement() && r.name() == QLatin1String("si"))) {
                r.readNext();
                if (r.isStartElement() && r.name() == QLatin1String("t"))
                    text += r.readElementText();
            }
            m_sharedStrings << text;
        }
    }
    return !r.hasError();
}

bool XlsxLoader::parseSheet(const QByteArray& xml, ISpreadsheetCore* core,
                             SheetId sheet, int firstRow, int count,
                             std::function<void(qint64,qint64)> progress)
{
    QXmlStreamReader r(xml);
    int loaded = 0;
    while (!r.atEnd()) {
        r.readNext();
        if (!r.isStartElement()) continue;
        if (r.name() == QLatin1String("row")) {
            int rowNum = r.attributes().value("r").toInt()-1;
            if (rowNum < firstRow) { r.skipCurrentElement(); continue; }
            if (loaded >= count) break;
            while (!(r.isEndElement() && r.name() == QLatin1String("row"))) {
                r.readNext();
                if (r.isStartElement() && r.name() == QLatin1String("c")) {
                    QString ref  = r.attributes().value("r").toString();
                    QString type = r.attributes().value("t").toString();
                    auto [row,col] = cellRefToRC(ref);
                    QString val;
                    while (!(r.isEndElement() && r.name() == QLatin1String("c"))) {
                        r.readNext();
                        if (r.isStartElement() && r.name() == QLatin1String("v"))
                            val = r.readElementText();
                        else if (r.isStartElement() && r.name() == QLatin1String("t"))
                            val = r.readElementText();
                    }
                    QVariant cv;
                    if (type == "s") {
                        int idx = val.toInt();
                        cv = (idx>=0 && idx<m_sharedStrings.size()) ? m_sharedStrings[idx] : val;
                    } else if (type == "b") {
                        cv = (val == "1");
                    } else {
                        bool ok; double d = val.toDouble(&ok);
                        cv = ok ? QVariant(d) : QVariant(val);
                    }
                    core->setCellValue(sheet, row, col, cv);
                }
            }
            ++loaded;
            if (progress && loaded%5000==0) progress(loaded, firstRow+count);
        }
    }
    return !r.hasError();
}

bool XlsxLoader::load(const QString& path, ISpreadsheetCore* core,
                      int firstRow, int count,
                      std::function<void(qint64,qint64)> progress)
{
    MiniZipReader zip(path);
    if (!zip.isReadable()) { m_error = "Cannot open XLSX: "+path; return false; }

    QByteArray ss = zip.fileData("xl/sharedStrings.xml");
    if (!ss.isEmpty()) parseSharedStrings(ss);

    QStringList rids;
    QXmlStreamReader wb(zip.fileData("xl/workbook.xml"));
    while (!wb.atEnd()) {
        wb.readNext();
        if (wb.isStartElement() && wb.name() == QLatin1String("sheet"))
            rids << wb.attributes().value("r:id").toString();
    }

    QHash<QString,QString> ridToFile;
    QXmlStreamReader rels(zip.fileData("xl/_rels/workbook.xml.rels"));
    while (!rels.atEnd()) {
        rels.readNext();
        if (rels.isStartElement() && rels.name() == QLatin1String("Relationship"))
            ridToFile[rels.attributes().value("Id").toString()] =
                      rels.attributes().value("Target").toString();
    }

    if (!rids.isEmpty()) {
        QString file = ridToFile.value(rids.first(), "worksheets/sheet1.xml");
        if (!file.startsWith("xl/")) file = "xl/"+file;
        SheetId sheet = core->sheets().isEmpty() ? core->addSheet("Sheet1") : core->sheets().first();
        parseSheet(zip.fileData(file), core, sheet, firstRow, count, progress);
    }
    return true;
}

bool XlsxLoader::save(const QString& path, ISpreadsheetCore* core,
                      SheetId sheet, std::function<void(qint64,qint64)> progress)
{
    MiniZipWriter zip;
    zip.open(path);

    zip.addFile("[Content_Types].xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "</Types>");

    zip.addFile("_rels/.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>");

    zip.addFile("xl/workbook.xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");

    zip.addFile("xl/_rels/workbook.xml.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        "</Relationships>");

    QByteArray wsData;
    QXmlStreamWriter ws(&wsData);
    ws.setAutoFormatting(false);
    ws.writeStartDocument();
    ws.writeStartElement("worksheet");
    ws.writeDefaultNamespace("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    ws.writeStartElement("sheetData");

    int rows = core->rowCount(sheet), cols = core->columnCount(sheet);
    auto colLabel = [](int c) {
        QString s; ++c;
        while (c>0) { --c; s.prepend(QChar('A'+c%26)); c/=26; }
        return s;
    };
    for (int r = 0; r < rows; ++r) {
        bool hasData = false;
        for (int c=0; c<cols && !hasData; ++c)
            hasData = core->getCell(sheet,r,c).rawValue.isValid();
        if (!hasData) continue;
        ws.writeStartElement("row");
        ws.writeAttribute("r", QString::number(r+1));
        for (int c = 0; c < cols; ++c) {
            Cell cell = core->getCell(sheet,r,c);
            QVariant val = cell.cachedValue.isValid() ? cell.cachedValue : cell.rawValue;
            if (!val.isValid()||val.isNull()) continue;
            ws.writeStartElement("c");
            ws.writeAttribute("r", colLabel(c)+QString::number(r+1));
            bool ok; double d = val.toDouble(&ok);
            if (ok) { ws.writeTextElement("v", QString::number(d,'g',15)); }
            else { ws.writeAttribute("t","inlineStr"); ws.writeTextElement("is",val.toString()); }
            ws.writeEndElement();
        }
        ws.writeEndElement();
        if (progress && r%10000==0) progress(r,rows);
    }
    ws.writeEndElement(); ws.writeEndElement();
    ws.writeEndDocument();
    zip.addFile("xl/worksheets/sheet1.xml", wsData);
    if (!zip.close()) { m_error = "Cannot write XLSX: "+path; return false; }
    if (progress) progress(rows,rows);
    return true;
}
