#include "OpsLoader.h"
#include "ISpreadsheetCore.h"
#include <QFile>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>

// ─────────────────────────────────────────────────────────────────────────────
//  Load
// ─────────────────────────────────────────────────────────────────────────────
bool OpsLoader::load(const QString& filePath, ISpreadsheetCore* target,
                     ProgressCallback progress)
{
    QFile f(filePath);
    if (!f.open(QIODevice::ReadOnly)) {
        m_error = f.errorString();
        return false;
    }
    const qint64 total = f.size();
    QByteArray data = f.readAll();
    f.close();
    if (progress) progress(total, total);

    QJsonParseError err;
    QJsonDocument doc = QJsonDocument::fromJson(data, &err);
    if (doc.isNull()) {
        m_error = "JSON parse error: " + err.errorString();
        return false;
    }
    QJsonObject root = doc.object();
    if (root["version"].toString() != "1.0") {
        m_error = "Unsupported .ops version";
        return false;
    }

    QJsonArray sheets = root["sheets"].toArray();
    for (const auto& sheetVal : sheets) {
        QJsonObject sh = sheetVal.toObject();
        QString name = sh["name"].toString("Sheet");
        SheetId sid = target->addSheet(name);

        // ── Cells ──────────────────────────────────────────────────────────
        QJsonArray cells = sh["cells"].toArray();
        for (const auto& cellVal : cells) {
            QJsonObject c = cellVal.toObject();
            int row = c["r"].toInt();
            int col = c["c"].toInt();
            QString formula = c["f"].toString();
            QString value   = c["v"].toString();
            if (!formula.isEmpty())
                target->setCellFormula(sid, row, col, formula);
            else
                target->setCellValue(sid, row, col, value);
        }

        // ── Row metadata ──────────────────────────────────────────────────
        QJsonArray rowMeta = sh["rowMeta"].toArray();
        for (const auto& rmVal : rowMeta) {
            QJsonObject rm = rmVal.toObject();
            RowMeta meta;
            meta.height = rm["h"].toInt(24);
            meta.hidden = rm["hidden"].toBool(false);
            target->setRowMeta(sid, rm["r"].toInt(), meta);
        }

        // ── Column metadata ───────────────────────────────────────────────
        QJsonArray colMeta = sh["colMeta"].toArray();
        for (const auto& cmVal : colMeta) {
            QJsonObject cm = cmVal.toObject();
            ColumnMeta meta;
            meta.width  = cm["w"].toInt(100);
            meta.hidden = cm["hidden"].toBool(false);
            target->setColMeta(sid, cm["c"].toInt(), meta);
        }
    }
    return true;
}

// ─────────────────────────────────────────────────────────────────────────────
//  Save
// ─────────────────────────────────────────────────────────────────────────────
bool OpsLoader::save(const QString& filePath, ISpreadsheetCore* source,
                     ProgressCallback progress)
{
    QJsonObject root;
    root["version"] = "1.0";

    QJsonArray sheets;
    for (SheetId sid : source->sheets()) {
        QJsonObject sh;
        sh["id"]   = sid;
        sh["name"] = source->sheetName(sid);

        int rows = source->rowCount(sid);
        int cols = source->columnCount(sid);

        // ── Cells (sparse: skip empty) ─────────────────────────────────────
        QJsonArray cells;
        for (int r = 0; r < rows; ++r) {
            for (int c = 0; c < cols; ++c) {
                Cell cell = source->getCell(sid, r, c);
                bool hasFormula = !cell.formula.isEmpty();
                bool hasValue   = !cell.rawValue.isNull() && !cell.rawValue.toString().isEmpty();
                if (!hasFormula && !hasValue) continue;
                QJsonObject jc;
                jc["r"] = r;
                jc["c"] = c;
                jc["v"] = cell.rawValue.toString();
                jc["f"] = cell.formula;
                cells.append(jc);
            }
        }
        sh["cells"] = cells;

        // ── Row metadata (only non-default rows) ──────────────────────────
        QJsonArray rowMeta;
        for (int r = 0; r < rows; ++r) {
            RowMeta rm = source->getRowMeta(sid, r);
            if (rm.height != 24 || rm.hidden) {
                QJsonObject jrm;
                jrm["r"]      = r;
                jrm["h"]      = rm.height;
                jrm["hidden"] = rm.hidden;
                rowMeta.append(jrm);
            }
        }
        sh["rowMeta"] = rowMeta;

        // ── Column metadata (only non-default cols) ───────────────────────
        QJsonArray colMeta;
        for (int c = 0; c < cols; ++c) {
            ColumnMeta cm = source->getColMeta(sid, c);
            if (cm.width != 100 || cm.hidden) {
                QJsonObject jcm;
                jcm["c"]      = c;
                jcm["w"]      = cm.width;
                jcm["hidden"] = cm.hidden;
                colMeta.append(jcm);
            }
        }
        sh["colMeta"] = colMeta;

        sheets.append(sh);
    }
    root["sheets"] = sheets;

    QByteArray data = QJsonDocument(root).toJson(QJsonDocument::Compact);
    QFile f(filePath);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        m_error = f.errorString();
        return false;
    }
    f.write(data);
    f.close();
    if (progress) progress(data.size(), data.size());
    return true;
}
