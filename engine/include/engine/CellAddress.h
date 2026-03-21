#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  CellAddress.h — Efficient cell coordinate type
//  Supports 1,048,576 rows × 16,384 columns (Excel limits)
// ═══════════════════════════════════════════════════════════════════════════════
#include <cstdint>
#include <functional>
#include <QString>
#include <QtGlobal>   // for quint32, size_t, Q_DECL_PURE_FUNCTION

static constexpr int MAX_ROWS = 1'048'576;
static constexpr int MAX_COLS = 16'384;

struct CellAddress {
    int row { 0 };
    int col { 0 };

    CellAddress() = default;
    CellAddress(int r, int c) : row(r), col(c) {}

    bool operator==(const CellAddress& o) const noexcept {
        return row == o.row && col == o.col;
    }
    bool operator!=(const CellAddress& o) const noexcept { return !(*this == o); }
    bool operator< (const CellAddress& o) const noexcept {
        return row != o.row ? row < o.row : col < o.col;
    }

    QString toRef() const {
        QString cs; int c = col + 1;
        while (c > 0) { --c; cs.prepend(QChar('A' + c % 26)); c /= 26; }
        return cs + QString::number(row + 1);
    }

    static CellAddress fromRef(const QString& ref) {
        QString upper = ref.toUpper().remove('$');
        int i = 0, c = 0;
        while (i < upper.size() && upper[i].isLetter())
            c = c * 26 + (upper[i++].unicode() - 'A' + 1);
        return { upper.mid(i).toInt() - 1, c - 1 };
    }

    bool isValid() const noexcept {
        return row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLS;
    }
};

// ── Qt hash support ───────────────────────────────────────────────────────────
// Implemented with plain arithmetic — no ::qHash() call, no Qt overload lookup.
// This avoids MSVC C2665 which fires when the compiler sees ::qHash(quint32)
// before <QHashFunctions> is fully resolved in the translation unit.
inline size_t qHash(const CellAddress& a, size_t seed = 0) noexcept {
    // Combine row and col with FNV-1a-inspired mixing; no Qt helpers needed.
    size_t h = seed;
    h ^= static_cast<size_t>(static_cast<unsigned int>(a.row)) + 0x9e3779b9u + (h << 6) + (h >> 2);
    h ^= static_cast<size_t>(static_cast<unsigned int>(a.col)) + 0x9e3779b9u + (h << 6) + (h >> 2);
    return h;
}

// ── std::hash for std::unordered_map ──────────────────────────────────────────
namespace std {
template<> struct hash<CellAddress> {
    size_t operator()(const CellAddress& a) const noexcept {
        size_t h = static_cast<size_t>(static_cast<unsigned int>(a.row)) * 1048583ULL
                 ^ static_cast<size_t>(static_cast<unsigned int>(a.col)) * 16777259ULL;
        return h ^ (h >> 16);
    }
};
}
