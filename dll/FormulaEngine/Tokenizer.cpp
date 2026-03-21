#include "Tokenizer.h"
#include <QRegularExpression>

Tokenizer::Tokenizer(const QString& expr) : m_expr(expr) {}

QString Tokenizer::colLabel(int col) {
    QString label;
    col++; // 1-based
    while (col > 0) {
        col--;
        label.prepend(QChar('A' + (col % 26)));
        col /= 26;
    }
    return label;
}

int Tokenizer::colIndex(const QString& col) {
    int idx = 0;
    for (QChar c : col.toUpper()) {
        idx = idx * 26 + (c.unicode() - 'A' + 1);
    }
    return idx - 1;
}

QList<Token> Tokenizer::tokenize() {
    QList<Token> tokens;
    m_pos = 0;
    while (m_pos < m_expr.size()) {
        QChar c = m_expr[m_pos];
        if (c.isSpace()) { m_pos++; continue; }
        if (c == '"')          tokens << readString();
        else if (c.isDigit() || (c == '.' && m_pos+1 < m_expr.size() && m_expr[m_pos+1].isDigit()))
                               tokens << readNumber();
        else if (c == '$' || c.isLetter())
                               tokens << readIdentifier();
        else if (c == '(')    { tokens << Token{TokenType::LParen,  "("}; m_pos++; }
        else if (c == ')')    { tokens << Token{TokenType::RParen,  ")"}; m_pos++; }
        else if (c == ',')    { tokens << Token{TokenType::Comma,   ","}; m_pos++; }
        else if (c == '+' || c == '-' || c == '*' || c == '/' ||
                 c == '^' || c == '&' || c == '%') {
            tokens << Token{TokenType::Operator, QString(c)};
            m_pos++;
        }
        else if (c == '<' || c == '>' || c == '=') {
            QString op(c); m_pos++;
            if (m_pos < m_expr.size() && (m_expr[m_pos]=='=' || m_expr[m_pos]=='>'))
                op += m_expr[m_pos++];
            tokens << Token{TokenType::Operator, op};
        }
        else m_pos++;
    }
    tokens << Token{TokenType::End, ""};
    return tokens;
}

Token Tokenizer::readNumber() {
    int start = m_pos;
    while (m_pos < m_expr.size() && (m_expr[m_pos].isDigit() || m_expr[m_pos] == '.'))
        m_pos++;
    // Scientific notation
    if (m_pos < m_expr.size() && (m_expr[m_pos] == 'e' || m_expr[m_pos] == 'E')) {
        m_pos++;
        if (m_pos < m_expr.size() && (m_expr[m_pos] == '+' || m_expr[m_pos] == '-'))
            m_pos++;
        while (m_pos < m_expr.size() && m_expr[m_pos].isDigit()) m_pos++;
    }
    QString s = m_expr.mid(start, m_pos - start);
    return Token{TokenType::Number, s, s.toDouble()};
}

Token Tokenizer::readString() {
    m_pos++; // skip opening "
    QString s;
    while (m_pos < m_expr.size()) {
        QChar c = m_expr[m_pos++];
        if (c == '"') {
            if (m_pos < m_expr.size() && m_expr[m_pos] == '"') { s += '"'; m_pos++; }
            else break;
        } else s += c;
    }
    return Token{TokenType::String, s};
}

Token Tokenizer::readIdentifier() {
    int start = m_pos;
    // skip $ signs and alphanumerics
    while (m_pos < m_expr.size() && (m_expr[m_pos] == '$' || m_expr[m_pos].isLetterOrNumber()))
        m_pos++;

    QString id = m_expr.mid(start, m_pos - start);

    // Check for range: A1:B10
    if (m_pos < m_expr.size() && m_expr[m_pos] == ':') {
        int saved = m_pos;
        m_pos++;
        int start2 = m_pos;
        while (m_pos < m_expr.size() && (m_expr[m_pos] == '$' || m_expr[m_pos].isLetterOrNumber()))
            m_pos++;
        if (m_pos > start2) {
            QString range = id + ":" + m_expr.mid(start2, m_pos - start2);
            return Token{TokenType::CellRange, range};
        }
        m_pos = saved; // backtrack
    }

    // Bool literals
    QString upper = id.toUpper().remove('$');
    if (upper == "TRUE")  return Token{TokenType::Bool,   "TRUE",  1.0};
    if (upper == "FALSE") return Token{TokenType::Bool,   "FALSE", 0.0};

    // Cell reference: letter(s) followed by digit(s)
    static QRegularExpression cellRe(R"(^\$?([A-Za-z]+)\$?(\d+)$)");
    if (cellRe.match(id).hasMatch())
        return Token{TokenType::CellRef, id.toUpper().remove('$')};

    // Otherwise function name
    return Token{TokenType::Function, id.toUpper()};
}
