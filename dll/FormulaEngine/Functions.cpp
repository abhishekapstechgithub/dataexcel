#include "Functions.h"
#include <cmath>
#include <algorithm>
#include <numeric>
#include <QtMath>
#include <QDateTime>
#include <QDate>

double FunctionRegistry::toNum(const QVariant& v) {
    if (v.typeId() == QMetaType::Bool) return v.toBool() ? 1.0 : 0.0;
    return v.toDouble();
}

QList<double> FunctionRegistry::toNums(const QList<QVariant>& args) {
    QList<double> out;
    for (auto& a : args) {
        if (!a.isNull() && a.isValid()) out << toNum(a);
    }
    return out;
}

void FunctionRegistry::registerFn(const QString& name, FnCallable fn) {
    m_fns[name.toUpper()] = fn;
}

bool FunctionRegistry::has(const QString& name) const {
    return m_fns.contains(name.toUpper());
}

QVariant FunctionRegistry::call(const QString& name, const QList<QVariant>& args) {
    auto it = m_fns.find(name.toUpper());
    if (it == m_fns.end()) return QString("#NAME?");
    try { return it.value()(args); }
    catch (...) { return QString("#ERROR!"); }
}

void FunctionRegistry::registerAll() {
    // ── Math ─────────────────────────────────────────────────────────────────
    registerFn("SUM", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        return std::accumulate(nums.begin(), nums.end(), 0.0);
    });
    registerFn("AVERAGE", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.isEmpty()) return QString("#DIV/0!");
        return std::accumulate(nums.begin(), nums.end(), 0.0) / nums.size();
    });
    registerFn("AVG", [this](QList<QVariant> a) -> QVariant {
        return call("AVERAGE", a);
    });
    registerFn("MIN", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.isEmpty()) return 0.0;
        return *std::min_element(nums.begin(), nums.end());
    });
    registerFn("MAX", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.isEmpty()) return 0.0;
        return *std::max_element(nums.begin(), nums.end());
    });
    registerFn("COUNT", [this](QList<QVariant> a) -> QVariant {
        return (double)toNums(a).size();
    });
    registerFn("COUNTA", [](QList<QVariant> a) -> QVariant {
        int cnt = 0;
        for (auto& v : a) if (!v.isNull() && v.isValid() && !v.toString().isEmpty()) cnt++;
        return (double)cnt;
    });
    registerFn("COUNTBLANK", [](QList<QVariant> a) -> QVariant {
        int cnt = 0;
        for (auto& v : a) if (v.isNull() || !v.isValid() || v.toString().isEmpty()) cnt++;
        return (double)cnt;
    });
    registerFn("ABS",  [this](QList<QVariant> a) -> QVariant { return std::abs(toNum(a.value(0))); });
    registerFn("SQRT", [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        if (v < 0) return QString("#NUM!");
        return std::sqrt(v);
    });
    registerFn("POWER", [this](QList<QVariant> a) -> QVariant {
        return std::pow(toNum(a.value(0)), toNum(a.value(1)));
    });
    registerFn("MOD",   [this](QList<QVariant> a) -> QVariant {
        double d = toNum(a.value(1));
        if (d == 0.0) return QString("#DIV/0!");
        return std::fmod(toNum(a.value(0)), d);
    });
    registerFn("INT",   [this](QList<QVariant> a) -> QVariant { return std::floor(toNum(a.value(0))); });
    registerFn("ROUND", [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        int d = (int)toNum(a.value(1));
        double factor = std::pow(10.0, d);
        return std::round(v * factor) / factor;
    });
    registerFn("ROUNDUP", [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        int d = (int)toNum(a.value(1));
        double f = std::pow(10.0, d);
        return std::ceil(v * f) / f;
    });
    registerFn("ROUNDDOWN", [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        int d = (int)toNum(a.value(1));
        double f = std::pow(10.0, d);
        return std::floor(v * f) / f;
    });
    registerFn("CEILING", [this](QList<QVariant> a) -> QVariant { return std::ceil(toNum(a.value(0))); });
    registerFn("FLOOR",   [this](QList<QVariant> a) -> QVariant { return std::floor(toNum(a.value(0))); });
    registerFn("TRUNC",   [this](QList<QVariant> a) -> QVariant { return std::trunc(toNum(a.value(0))); });
    registerFn("EXP",  [this](QList<QVariant> a) -> QVariant { return std::exp(toNum(a.value(0))); });
    registerFn("LN",   [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        if (v <= 0) return QString("#NUM!");
        return std::log(v);
    });
    registerFn("LOG10",[this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        if (v <= 0) return QString("#NUM!");
        return std::log10(v);
    });
    registerFn("LOG",  [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        double base = a.size() > 1 ? toNum(a.value(1)) : 10.0;
        if (v <= 0 || base <= 0 || base == 1) return QString("#NUM!");
        return std::log(v) / std::log(base);
    });
    registerFn("PI",   [](QList<QVariant>) -> QVariant { return M_PI; });
    registerFn("SIN",  [this](QList<QVariant> a) -> QVariant { return std::sin(toNum(a.value(0))); });
    registerFn("COS",  [this](QList<QVariant> a) -> QVariant { return std::cos(toNum(a.value(0))); });
    registerFn("TAN",  [this](QList<QVariant> a) -> QVariant { return std::tan(toNum(a.value(0))); });
    registerFn("RAND", [](QList<QVariant>) -> QVariant { return (double)rand() / RAND_MAX; });
    registerFn("RANDBETWEEN", [this](QList<QVariant> a) -> QVariant {
        int lo = (int)toNum(a.value(0)), hi = (int)toNum(a.value(1));
        return (double)(lo + rand() % (hi - lo + 1));
    });

    // ── Statistical ───────────────────────────────────────────────────────────
    registerFn("STDEV", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.size() < 2) return QString("#DIV/0!");
        double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / nums.size();
        double sq = 0;
        for (double v : nums) sq += (v - mean) * (v - mean);
        return std::sqrt(sq / (nums.size() - 1));
    });
    registerFn("VAR", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.size() < 2) return QString("#DIV/0!");
        double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / nums.size();
        double sq = 0;
        for (double v : nums) sq += (v - mean) * (v - mean);
        return sq / (nums.size() - 1);
    });
    registerFn("MEDIAN", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.isEmpty()) return 0.0;
        std::sort(nums.begin(), nums.end());
        int n = nums.size();
        return (n % 2 == 0) ? (nums[n/2-1] + nums[n/2]) / 2.0 : nums[n/2];
    });
    registerFn("LARGE", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        int k = (int)toNum(a.value(a.size()-1));
        std::sort(nums.begin(), nums.end(), std::greater<double>());
        if (k < 1 || k > nums.size()) return QString("#NUM!");
        return nums[k-1];
    });
    registerFn("SMALL", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        int k = (int)toNum(a.value(a.size()-1));
        std::sort(nums.begin(), nums.end());
        if (k < 1 || k > nums.size()) return QString("#NUM!");
        return nums[k-1];
    });

    // ── Logical ───────────────────────────────────────────────────────────────
    registerFn("IF", [this](QList<QVariant> a) -> QVariant {
        bool cond = a.value(0).toBool();
        if (a.value(0).typeId() == QMetaType::Double) cond = (toNum(a.value(0)) != 0.0);
        return cond ? a.value(1) : a.value(2);
    });
    registerFn("AND", [this](QList<QVariant> a) -> QVariant {
        for (auto& v : a) if (!v.toBool() && toNum(v) == 0.0) return false;
        return true;
    });
    registerFn("OR", [this](QList<QVariant> a) -> QVariant {
        for (auto& v : a) if (v.toBool() || toNum(v) != 0.0) return true;
        return false;
    });
    registerFn("NOT", [this](QList<QVariant> a) -> QVariant {
        return !(a.value(0).toBool() || toNum(a.value(0)) != 0.0);
    });
    registerFn("IFERROR", [](QList<QVariant> a) -> QVariant {
        QString s = a.value(0).toString();
        if (s.startsWith('#')) return a.value(1);
        return a.value(0);
    });
    registerFn("ISBLANK",  [](QList<QVariant> a) -> QVariant { return a.value(0).isNull() || a.value(0).toString().isEmpty(); });
    registerFn("ISNUMBER", [](QList<QVariant> a) -> QVariant {
        bool ok; a.value(0).toDouble(&ok); return ok;
    });
    registerFn("ISTEXT",   [](QList<QVariant> a) -> QVariant {
        bool ok; a.value(0).toDouble(&ok); return !ok && !a.value(0).isNull();
    });

    // ── String ────────────────────────────────────────────────────────────────
    registerFn("LEN",     [](QList<QVariant> a) -> QVariant { return (double)a.value(0).toString().length(); });
    registerFn("LEFT",    [this](QList<QVariant> a) -> QVariant { return a.value(0).toString().left((int)toNum(a.value(1))); });
    registerFn("RIGHT",   [this](QList<QVariant> a) -> QVariant { return a.value(0).toString().right((int)toNum(a.value(1))); });
    registerFn("MID",     [this](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().mid((int)toNum(a.value(1))-1, (int)toNum(a.value(2)));
    });
    registerFn("UPPER",   [](QList<QVariant> a) -> QVariant { return a.value(0).toString().toUpper(); });
    registerFn("LOWER",   [](QList<QVariant> a) -> QVariant { return a.value(0).toString().toLower(); });
    registerFn("TRIM",    [](QList<QVariant> a) -> QVariant { return a.value(0).toString().trimmed(); });
    registerFn("CONCATENATE", [](QList<QVariant> a) -> QVariant {
        QString s;
        for (auto& v : a) s += v.toString();
        return s;
    });
    registerFn("SUBSTITUTE", [](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().replace(a.value(1).toString(), a.value(2).toString());
    });
    registerFn("FIND", [](QList<QVariant> a) -> QVariant {
        int idx = a.value(1).toString().indexOf(a.value(0).toString(),
                  a.size() > 2 ? (int)a.value(2).toDouble()-1 : 0);
        return idx < 0 ? QVariant(QString("#VALUE!")) : QVariant((double)(idx+1));
    });
    registerFn("VALUE",  [this](QList<QVariant> a) -> QVariant { return toNum(a.value(0)); });
    registerFn("TEXT",   [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        QString fmt = a.value(1).toString();
        if (fmt.contains('%')) return QString::number(v * 100, 'f', 1) + "%";
        if (fmt.contains("0.00")) return QString::number(v, 'f', 2);
        return QString::number(v);
    });
    registerFn("REPT", [this](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().repeated((int)toNum(a.value(1)));
    });

    // ── Lookup ────────────────────────────────────────────────────────────────
    registerFn("VLOOKUP", [this](QList<QVariant> a) -> QVariant {
        // VLOOKUP(lookup_val, range_values..., col_index, [exact])
        // Simplified: args = [lookup, v1, v2, ..., vN, col_index, match_type]
        // For real impl a range would be passed as a flat list
        if (a.size() < 3) return QString("#N/A");
        QVariant target = a.value(0);
        int colIdx = (int)toNum(a.last()) - 1;
        // rest are rows flattened - simplified to just return #N/A for now
        return QString("#N/A");
    });

    // ── Date (basic) ─────────────────────────────────────────────────────────
    registerFn("NOW",   [](QList<QVariant>) -> QVariant { return QDateTime::currentDateTime().toString(); });
    registerFn("TODAY", [](QList<QVariant>) -> QVariant { return QDate::currentDate().toString(); });
    registerFn("YEAR",  [](QList<QVariant> a) -> QVariant {
        return (double)QDate::fromString(a.value(0).toString(), Qt::ISODate).year();
    });
    registerFn("MONTH", [](QList<QVariant> a) -> QVariant {
        return (double)QDate::fromString(a.value(0).toString(), Qt::ISODate).month();
    });
    registerFn("DAY",   [](QList<QVariant> a) -> QVariant {
        return (double)QDate::fromString(a.value(0).toString(), Qt::ISODate).day();
    });
}
