#include "Parser.h"
#include "Functions.h"
#include <stdexcept>
#include <cmath>
#include <QRegularExpression>

Parser::Parser(const QList<Token>& tokens, CellResolver resolver, FunctionRegistry* registry)
    : m_tokens(tokens), m_resolver(resolver), m_registry(registry) {}

Token& Parser::cur()  { return m_tokens[m_pos]; }
Token& Parser::peek(int o) {
    int idx = m_pos + o;
    if (idx >= m_tokens.size()) return m_tokens.last();
    return m_tokens[idx];
}
Token Parser::consume() { return m_tokens[m_pos++]; }
bool Parser::match(TokenType t) {
    if (cur().type == t) { m_pos++; return true; }
    return false;
}

QVariant Parser::parse() { return parseExpr(); }
QVariant Parser::parseExpr() { return parseComparison(); }

QVariant Parser::parseComparison() {
    auto left = parseAddSub();
    while (cur().type == TokenType::Operator &&
           (cur().value == "=" || cur().value == "<>" ||
            cur().value == "<" || cur().value == ">" ||
            cur().value == "<=" || cur().value == ">=")) {
        QString op = consume().value;
        auto right = parseAddSub();
        double l = toDouble(left), r = toDouble(right);
        if (op == "=")       left = (l == r);
        else if (op == "<>") left = (l != r);
        else if (op == "<")  left = (l < r);
        else if (op == ">")  left = (l > r);
        else if (op == "<=") left = (l <= r);
        else if (op == ">=") left = (l >= r);
    }
    while (cur().type == TokenType::Operator && cur().value == "&") {
        consume();
        auto right = parseAddSub();
        left = toString(left) + toString(right);
    }
    return left;
}

QVariant Parser::parseAddSub() {
    auto left = parseMulDiv();
    while (cur().type == TokenType::Operator &&
           (cur().value == "+" || cur().value == "-")) {
        QString op = consume().value;
        auto right = parseMulDiv();
        if (op == "+") left = toDouble(left) + toDouble(right);
        else           left = toDouble(left) - toDouble(right);
    }
    return left;
}

QVariant Parser::parseMulDiv() {
    auto left = parsePow();
    while (cur().type == TokenType::Operator &&
           (cur().value == "*" || cur().value == "/")) {
        QString op = consume().value;
        auto right = parsePow();
        double r = toDouble(right);
        if (op == "*") left = toDouble(left) * r;
        else {
            if (r == 0.0) return QString("#DIV/0!");
            left = toDouble(left) / r;
        }
    }
    return left;
}

QVariant Parser::parsePow() {
    auto base = parseUnary();
    if (cur().type == TokenType::Operator && cur().value == "^") {
        consume();
        return std::pow(toDouble(base), toDouble(parseUnary()));
    }
    return base;
}

QVariant Parser::parseUnary() {
    if (cur().type == TokenType::Operator && cur().value == "-") {
        consume(); return -toDouble(parseAtom());
    }
    if (cur().type == TokenType::Operator && cur().value == "+") {
        consume(); return toDouble(parseAtom());
    }
    return parseAtom();
}

QVariant Parser::parseAtom() {
    Token tok = consume();
    switch (tok.type) {
    case TokenType::Number:  return tok.number;
    case TokenType::String:  return tok.value;
    case TokenType::Bool:    return (tok.number != 0.0);
    case TokenType::CellRef:
        if (m_resolver) return m_resolver(tok.value);
        return QVariant();
    case TokenType::CellRange: return QVariant();
    case TokenType::Function: {
        QString fname = tok.value;
        if (cur().type != TokenType::LParen)
            throw std::runtime_error(("Expected '(' after: " + fname).toStdString());
        consume();
        QList<QVariant> args;
        while (cur().type != TokenType::RParen && cur().type != TokenType::End) {
            if (cur().type == TokenType::CellRange) {
                // Pass as a single structured QVariant (list-of-rows).
                // toNums() will flatten it for aggregate functions;
                // 2D-aware functions (VLOOKUP, SUMIF…) read the structure directly.
                args << expandRange2D(consume().value);
            } else {
                args << parseExpr();
            }
            if (cur().type == TokenType::Comma) consume();
        }
        match(TokenType::RParen);
        return callFunction(fname, args);
    }
    case TokenType::LParen: {
        auto val = parseExpr();
        match(TokenType::RParen);
        return val;
    }
    default: return QVariant();
    }
}

QList<QVariant> Parser::expandRange(const QString& range) {
    QList<QVariant> values;
    static QRegularExpression re(R"(([A-Z]+)(\d+):([A-Z]+)(\d+))");
    auto m = re.match(range);
    if (!m.hasMatch()) return values;
    int c1 = Tokenizer::colIndex(m.captured(1));
    int r1 = m.captured(2).toInt();
    int c2 = Tokenizer::colIndex(m.captured(3));
    int r2 = m.captured(4).toInt();
    for (int r = r1; r <= r2; ++r)
        for (int c = c1; c <= c2; ++c)
            if (m_resolver)
                values << m_resolver(Tokenizer::colLabel(c) + QString::number(r));
    return values;
}

// Returns a QVariant wrapping a QList<QVariant> where each element is a
// QVariant wrapping a QList<QVariant> (one per row).  Functions that only
// need a flat list of numbers call toNums() which recursively flattens it;
// 2D-aware functions (VLOOKUP, SUMIF …) read the row/column structure directly.
QVariant Parser::expandRange2D(const QString& range) {
    static QRegularExpression re(R"(([A-Z]+)(\d+):([A-Z]+)(\d+))");
    auto m = re.match(range);
    if (!m.hasMatch()) return QVariant(expandRange(range)); // fallback: flat
    int c1 = Tokenizer::colIndex(m.captured(1));
    int r1 = m.captured(2).toInt();
    int c2 = Tokenizer::colIndex(m.captured(3));
    int r2 = m.captured(4).toInt();
    QList<QVariant> rows;
    rows.reserve(r2 - r1 + 1);
    for (int r = r1; r <= r2; ++r) {
        QList<QVariant> row;
        row.reserve(c2 - c1 + 1);
        for (int c = c1; c <= c2; ++c)
            if (m_resolver)
                row << m_resolver(Tokenizer::colLabel(c) + QString::number(r));
        rows << QVariant(row);
    }
    return QVariant(rows);
}

QVariant Parser::callFunction(const QString& name, QList<QVariant> args) {
    if (!m_registry) return QString("#NAME?");
    return m_registry->call(name, args);
}

double Parser::toDouble(const QVariant& v) {
    if (v.typeId() == QMetaType::Bool)   return v.toBool() ? 1.0 : 0.0;
    if (v.typeId() == QMetaType::QString) return v.toString().toDouble();
    return v.toDouble();
}

bool Parser::toBool(const QVariant& v) {
    if (v.typeId() == QMetaType::Bool) return v.toBool();
    return toDouble(v) != 0.0;
}

QString Parser::toString(const QVariant& v) {
    if (!v.isValid() || v.isNull()) return QString();
    return v.toString();
}
