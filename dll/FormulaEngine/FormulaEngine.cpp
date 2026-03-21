#include "FormulaEngine.h"
#include "Tokenizer.h"
#include "Parser.h"
#include "Functions.h"
#include <QStack>
#include <QRegularExpression>
#include <cmath>

// ── FormulaEngine::Impl ──────────────────────────────────────────────────────
struct FormulaEngine::Impl {
    QString lastError;
    FunctionRegistry registry;

    Impl() { registry.registerAll(); }

    QVariant evaluate(const QString& formula, CellResolver resolver);
    QSet<CellRef> dependencies(const QString& formula);
};

// ── Public API ───────────────────────────────────────────────────────────────
FormulaEngine::FormulaEngine() : d(new Impl()) {}
FormulaEngine::~FormulaEngine() { delete d; }

QVariant FormulaEngine::evaluate(const QString& formula, CellResolver resolver) {
    d->lastError.clear();
    if (formula.isEmpty()) return QVariant();
    QString f = formula.startsWith('=') ? formula.mid(1).trimmed() : formula.trimmed();
    try {
        return d->evaluate(f, resolver);
    } catch (const std::exception& e) {
        d->lastError = QString::fromStdString(e.what());
        return QString("#ERROR: ") + d->lastError;
    }
}

QSet<CellRef> FormulaEngine::dependencies(const QString& formula) {
    return d->dependencies(formula);
}

QString FormulaEngine::lastError() const { return d->lastError; }

// ── Impl::evaluate ────────────────────────────────────────────────────────────
QVariant FormulaEngine::Impl::evaluate(const QString& formula, CellResolver resolver) {
    Tokenizer tokenizer(formula);
    auto tokens = tokenizer.tokenize();
    Parser parser(tokens, resolver, &registry);
    return parser.parse();
}

// ── Impl::dependencies ───────────────────────────────────────────────────────
QSet<CellRef> FormulaEngine::Impl::dependencies(const QString& formula) {
    QSet<CellRef> deps;
    if (!formula.startsWith('=')) return deps;
    QString f = formula.mid(1);

    // Match cell refs: A1, $A$1, A1:B10, $A$1:$B$10
    static QRegularExpression cellRe(
        R"(\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?)"
    );
    auto it = cellRe.globalMatch(f);
    while (it.hasNext()) {
        auto match = it.next();
        QString colStr1 = match.captured(1).toUpper();
        QString rowStr1 = match.captured(2);
        QString colStr2 = match.captured(3).toUpper();
        QString rowStr2 = match.captured(4);

        if (colStr2.isEmpty()) {
            // Single cell
            deps.insert(colStr1 + rowStr1);
        } else {
            // Range: expand
            int c1 = Tokenizer::colIndex(colStr1);
            int c2 = Tokenizer::colIndex(colStr2);
            int r1 = rowStr1.toInt();
            int r2 = rowStr2.toInt();
            for (int r = r1; r <= r2; ++r)
                for (int c = c1; c <= c2; ++c)
                    deps.insert(Tokenizer::colLabel(c) + QString::number(r));
        }
    }
    return deps;
}
