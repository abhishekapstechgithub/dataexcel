#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  CellAddress.h — Efficient cell coordinate type
//  Supports 1,048,576 rows × 16,384 columns (Excel limits)
// ═══════════════════════════════════════════════════════════════════════════════
#include <cstdint>
#include <functional>
#include <QString>

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

    // Pack into uint64 — used only by std::hash below
    std::uint64_t key() const noexcept {
        return (static_cast<std::uint64_t>(static_cast<unsigned>(row)) << 20)
             |  static_cast<std::uint64_t>(static_cast<unsigned>(col));
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

// ── Qt hash — split key into two quint32 to avoid int64 overload issues ────────
inline size_t qHash(const CellAddress& a, size_t seed = 0) noexcept {
    // XOR fold: combine row and col without touching int64 qHash
    quint32 packed = (static_cast<quint32>(a.row) << 14)
                   ^  static_cast<quint32>(a.col);
    return ::qHash(packed, seed);
}

// ── std::hash for std::unordered_map ──────────────────────────────────────────
namespace std {
template<> struct hash<CellAddress> {
    size_t operator()(const CellAddress& a) const noexcept {
        // FNV-inspired mix — avoids any platform-specific qHash
        size_t h = static_cast<size_t>(a.row) * 1048583ULL
                 ^ static_cast<size_t>(a.col) * 16777259ULL;
        return h ^ (h >> 16);
    }
};
}
