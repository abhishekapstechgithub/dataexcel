#pragma once
#include <QString>
#include <QList>

enum class TokenType {
    Number, String, CellRef, CellRange, Operator,
    Comma, LParen, RParen, Function, Bool, Error, End
};

struct Token {
    TokenType type;
    QString   value;
    double    number { 0.0 };
};

class Tokenizer {
public:
    explicit Tokenizer(const QString& expr);
    QList<Token> tokenize();

    static QString colLabel(int col);   // 0→A, 25→Z, 26→AA
    static int     colIndex(const QString& col); // "A"→0
private:
    QString m_expr;
    int     m_pos { 0 };
    Token   readNext();
    Token   readNumber();
    Token   readString();
    Token   readIdentifier();
};
