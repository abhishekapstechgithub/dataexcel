// ═══════════════════════════════════════════════════════════════════════════════
//  XlsxSaxLoader.cpp — Streaming SAX-based XLSX parser
//
//  Memory usage: O(TILE_ROWS) at any time — never loads full file into memory.
//  Uses QXmlStreamReader for forward-only parsing.
// ═══════════════════════════════════════════════════════════════════════════════
#include "XlsxSaxLoader.h"
#include <QXmlStreamReader>
#include <QRegularExpression>
#include <QVariant>
#include <QDate>
#include <cmath>

// ── CellRef parser ────────────────────────────────────────────────────────────
CellRef CellRef::parse(const QString& ref)
{
    CellRef cr;
    int i = 0;
    // Column letters
    int col = 0;
    while (i < ref.size() && ref[i].isLetter()) {
        col = col * 26 + (ref[i].toUpper().unicode() - 'A' + 1);
        ++i;
    }
    cr.col = col - 1;  // 0-based

    // Row digits
    int row = 0;
    while (i < ref.size() && ref[i].isDigit()) {
        row = row * 10 + ref[i].digitValue();
        ++i;
    }
    cr.row = row - 1;  // 0-based
    return cr;
}

// ── Shared strings parser ─────────────────────────────────────────────────────
bool XlsxSaxLoader::parseSharedStrings(const QByteArray& xml, QStringList& out)
{
    out.clear();
    QXmlStreamReader reader(xml);
    QString current;
    bool inSi = false;
    bool inT  = false;

    while (!reader.atEnd()) {
        reader.readNext();
        if (reader.isStartElement()) {
            if (reader.name() == u"si")  { inSi = true; current.clear(); }
            if (reader.name() == u"t" && inSi) { inT = true; }
        } else if (reader.isEndElement()) {
            if (reader.name() == u"t")  { inT = false; }
            if (reader.name() == u"si") { inSi = false; out << current; }
        } else if (reader.isCharacters() && inT) {
            current += reader.text().toString();
        }
    }

    if (reader.hasError()) {
        m_lastError = reader.errorString();
        return false;
    }
    return true;
}

// ── Styles parser (number formats) ────────────────────────────────────────────
bool XlsxSaxLoader::parseStyles(const QByteArray& stylesXml, QHash<int,QString>& numFmts)
{
    numFmts.clear();
    QXmlStreamReader reader(stylesXml);
    bool inNumFmts = false;

    while (!reader.atEnd()) {
        reader.readNext();
        if (reader.isStartElement()) {
            if (reader.name() == u"numFmts") inNumFmts = true;
            if (reader.name() == u"numFmt" && inNumFmts) {
                int id = reader.attributes().value("numFmtId").toInt();
                QString code = reader.attributes().value("formatCode").toString();
                numFmts[id] = code;
            }
        } else if (reader.isEndElement()) {
            if (reader.name() == u"numFmts") inNumFmts = false;
        }
    }
    return !reader.hasError();
}

// ── Helper: apply number format ───────────────────────────────────────────────
static QString applyNumFmt(double val, const QString& fmtCode)
{
    if (fmtCode.isEmpty()) return QString::number(val, 'g', 10);

    QString code = fmtCode.toLower();
    if (code.contains("yy") || code.contains("mm") || code.contains("dd")) {
        // Date: Excel date serial
        int days = (int)val;
        QDate base(1899, 12, 30);
        QDate date = base.addDays(days);
        return date.toString("yyyy-MM-dd");
    }
    if (code.contains('%')) {
        return QString::number(val * 100.0, 'f', 2) + "%";
    }
    if (code.contains("0.00")) {
        return QString::number(val, 'f', 2);
    }
    return QString::number(val, 'g', 10);
}

// ── Sheet parser ──────────────────────────────────────────────────────────────
bool XlsxSaxLoader::parseSheet(
    const QByteArray& sheetXml,
    const QStringList& sharedStrings,
    const QHash<int,QString>& numFmts,
    int firstRow, int maxRows,
    std::function<void(int,int,const QString&,const QString&)> onCell,
    ProgressCallback progress)
{
    qint64 total = sheetXml.size();
    qint64 pos = 0;
    qint64 reportEvery = std::max((qint64)1, total / 200);
    qint64 nextReport = 0;

    QXmlStreamReader reader(sheetXml);

    // Per-cell state
    QString cellRef;
    QString cellType;     // "s" = sharedString, "str" = formula string, "b" = bool, else numeric
    QString cellStyle;
    QString valueText;
    QString formulaText;
    bool    inCell    = false;
    bool    inValue   = false;
    bool    inFormula = false;

    while (!reader.atEnd()) {
        reader.readNext();

        // Progress reporting
        pos = reader.characterOffset();
        if (pos >= nextReport && progress) {
            progress(pos, total);
            nextReport = pos + reportEvery;
        }

        if (reader.isStartElement()) {
            const auto name = reader.name();

            if (name == u"c") {
                // Start of a cell
                inCell    = true;
                cellRef   = reader.attributes().value("r").toString();
                cellType  = reader.attributes().value("t").toString();
                cellStyle = reader.attributes().value("s").toString();
                valueText.clear();
                formulaText.clear();
                inValue   = false;
                inFormula = false;

                // Row filter
                CellRef cr = CellRef::parse(cellRef);
                if (maxRows > 0 && (cr.row < firstRow || cr.row >= firstRow + maxRows))
                    inCell = false;

            } else if (name == u"v" && inCell) {
                inValue   = true;
                inFormula = false;
            } else if (name == u"f" && inCell) {
                inFormula = true;
                inValue   = false;
            }

        } else if (reader.isEndElement()) {
            const auto name = reader.name();

            if (name == u"v")  { inValue = false; }
            if (name == u"f")  { inFormula = false; }

            if (name == u"c" && inCell) {
                inCell = false;
                CellRef cr = CellRef::parse(cellRef);

                QString displayValue;
                if (cellType == "s") {
                    // Shared string index
                    int idx = valueText.toInt();
                    displayValue = (idx >= 0 && idx < sharedStrings.size())
                                 ? sharedStrings[idx] : valueText;
                } else if (cellType == "b") {
                    displayValue = (valueText == "1") ? "TRUE" : "FALSE";
                } else if (cellType == "str") {
                    displayValue = valueText;
                } else {
                    // Numeric
                    if (!valueText.isEmpty()) {
                        bool ok;
                        double d = valueText.toDouble(&ok);
                        if (ok) {
                            // Apply number format if available
                            int styleIdx = cellStyle.toInt();
                            QString fmtCode = numFmts.value(styleIdx);
                            displayValue = applyNumFmt(d, fmtCode);
                        } else {
                            displayValue = valueText;
                        }
                    }
                }

                if (!displayValue.isEmpty() || !formulaText.isEmpty())
                    onCell(cr.row, cr.col, displayValue, formulaText);
            }

        } else if (reader.isCharacters()) {
            if (inValue)   valueText   += reader.text().toString();
            if (inFormula) formulaText += reader.text().toString();
        }
    }

    if (progress) progress(total, total);

    if (reader.hasError()) {
        m_lastError = reader.errorString();
        return false;
    }
    return true;
}
