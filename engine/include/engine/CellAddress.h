#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  CellAddress.h — Efficient cell coordinate type
//  Supports 1,048,576 rows × 16,384 columns (Excel limits)
//  Uses a single int64 for O(1) hash and comparison.
// ═══════════════════════════════════════════════════════════════════════════════
#include <cstdint>
#include <functional>
#include <QString>

// Excel maximums
static constexpr int MAX_ROWS = 1'048'576;
static constexpr int MAX_COLS = 16'384;

struct CellAddress {
    int row { 0 };   // 0-based
    int col { 0 };   // 0-based

    CellAddress() = default;
    CellAddress(int r, int c) : row(r), col(c) {}

    bool operator==(const CellAddress& o) const noexcept {
        return row == o.row && col == o.col;
    }
    bool operator!=(const CellAddress& o) const noexcept {
        return !(*this == o);
    }
    bool operator<(const CellAddress& o) const noexcept {
        return key() < o.key();
    }

    // Pack into int64 for fast hashing
    std::int64_t key() const noexcept {
        return (static_cast<std::int64_t>(row) << 20) | col;
    }

    // Convert to Excel-style "A1" reference (1-based display)
    QString toRef() const {
        QString col_str;
        int c = col + 1;
        while (c > 0) { --c; col_str.prepend(QChar('A' + c % 26)); c /= 26; }
        return col_str + QString::number(row + 1);
    }

    // Parse "A1", "$A$1", "AB12" etc.
    static CellAddress fromRef(const QString& ref) {
        QString upper = ref.toUpper().remove('$');
        int i = 0;
        int c = 0;
        while (i < upper.size() && upper[i].isLetter()) {
            c = c * 26 + (upper[i].unicode() - 'A' + 1);
            ++i;
        }
        int r = upper.mid(i).toInt();
        return { r - 1, c - 1 };
    }

    bool isValid() const noexcept {
        return row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLS;
    }
};

// Qt hash support
inline size_t qHash(const CellAddress& a, size_t seed = 0) noexcept {
    return qHash(a.key(), seed);
}

// std hash support (for unordered_map)
namespace std {
template<> struct hash<CellAddress> {
    size_t operator()(const CellAddress& a) const noexcept {
        return std::hash<std::int64_t>{}(a.key());
    }
};
}
