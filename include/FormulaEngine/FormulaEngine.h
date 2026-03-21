#pragma once
#ifdef _WIN32
  #ifdef FORMULAENGINE_EXPORTS
    #define FORMULA_API __declspec(dllexport)
  #else
    #define FORMULA_API __declspec(dllimport)
  #endif
#else
  #define FORMULA_API __attribute__((visibility("default")))
#endif

#include <QString>
#include <QVariant>
#include <QHash>
#include <QSet>
#include <functional>

// Cell reference type: "A1", "B3", etc.
using CellRef   = QString;
using CellValue = QVariant;

// Callback to resolve a cell value during formula evaluation
using CellResolver = std::function<CellValue(const CellRef&)>;

class FORMULA_API FormulaEngine
{
public:
    FormulaEngine();
    ~FormulaEngine();

    // Evaluate a formula string. resolver is called for each cell reference found.
    // Returns the computed result or an error string prefixed with "#"
    QVariant evaluate(const QString& formula, CellResolver resolver);

    // Returns all cell references found in a formula (for dependency tracking)
    QSet<CellRef> dependencies(const QString& formula);

    // Returns last error description (empty if no error)
    QString lastError() const;

private:
    struct Impl;
    Impl* d;
};
