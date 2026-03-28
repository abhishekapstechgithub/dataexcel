#pragma once
#include "Tokenizer.h"
#include "FormulaEngine.h"
#include <QVariant>

class FunctionRegistry;

class Parser {
public:
    Parser(const QList<Token>& tokens,
           CellResolver resolver,
           FunctionRegistry* registry);

    QVariant parse();

private:
    QList<Token>      m_tokens;
    int               m_pos { 0 };
    CellResolver      m_resolver;
    FunctionRegistry* m_registry;

    Token&   cur();
    Token&   peek(int offset = 1);
    Token    consume();
    bool     match(TokenType t);

    QVariant parseExpr();
    QVariant parseComparison();
    QVariant parseAddSub();
    QVariant parseMulDiv();
    QVariant parsePow();
    QVariant parseUnary();
    QVariant parseAtom();

    QVariant callFunction(const QString& name, QList<QVariant> args);
    QList<QVariant> expandRange(const QString& range);
    // Returns a single QVariant wrapping a list-of-rows (each row is a QVariantList).
    // Used so 2D-aware functions (VLOOKUP, SUMIF, etc.) can see the table structure.
    QVariant        expandRange2D(const QString& range);
    double   toDouble(const QVariant& v);
    bool     toBool(const QVariant& v);
    QString  toString(const QVariant& v);
};
