#pragma once
// ─────────────────────────────────────────────────────────────────────────────
//  OpsLoader — native OpenSheet workbook format (.ops)
//
//  Format: JSON file (optionally ZIP-compressed in future).
//  Schema:
//    {
//      "version": "1.0",
//      "sheets": [
//        {
//          "id": 0,
//          "name": "Sheet1",
//          "cells": [
//            { "r":0, "c":0, "v":"Hello", "f":"" },
//            { "r":1, "c":2, "v":"42",    "f":"=SUM(A1:A5)" }
//          ],
//          "rowMeta": [ { "r":0, "h":30, "hidden":false } ],
//          "colMeta": [ { "c":0, "w":120, "hidden":false } ]
//        }
//      ]
//    }
// ─────────────────────────────────────────────────────────────────────────────
#include <QString>
#include <functional>

class ISpreadsheetCore;
using ProgressCallback = std::function<void(qint64, qint64)>;

class OpsLoader
{
public:
    OpsLoader() = default;

    // Load all sheets from an .ops file into *target*.
    bool load(const QString& filePath,
              ISpreadsheetCore* target,
              ProgressCallback progress = {});

    // Save all sheets from *source* to an .ops file.
    bool save(const QString& filePath,
              ISpreadsheetCore* source,
              ProgressCallback progress = {});

    QString lastError() const { return m_error; }

private:
    QString m_error;
};
