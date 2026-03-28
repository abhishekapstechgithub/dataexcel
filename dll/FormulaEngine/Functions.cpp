#include "Functions.h"
#include <cmath>
#include <algorithm>
#include <numeric>
#include <QtMath>
#include <QDateTime>
#include <QDate>
#include <QRegularExpression>

double FunctionRegistry::toNum(const QVariant& v) {
    if (v.typeId() == QMetaType::Bool) return v.toBool() ? 1.0 : 0.0;
    return v.toDouble();
}

// Recursively flatten a QVariant that may be a list-of-rows into doubles.
static void flattenVariantToNums(const QVariant& v, QList<double>& out) {
    if (v.typeId() == QMetaType::QVariantList) {
        for (const auto& item : v.toList())
            flattenVariantToNums(item, out);
    } else if (!v.isNull() && v.isValid()) {
        bool ok;
        double d = v.toDouble(&ok);
        if (ok) out << d;
    }
}

QList<double> FunctionRegistry::toNums(const QList<QVariant>& args) {
    QList<double> out;
    for (const auto& a : args)
        flattenVariantToNums(a, out);
    return out;
}

// Flatten a QVariant (possibly list-of-rows) to a flat QList<QVariant>.
QList<QVariant> FunctionRegistry::flattenToList(const QVariant& v) {
    QList<QVariant> out;
    if (v.typeId() == QMetaType::QVariantList) {
        for (const auto& item : v.toList()) {
            auto sub = flattenToList(item);
            out.append(sub);
        }
    } else {
        out << v;
    }
    return out;
}

// Match a single value against an Excel-style criteria string.
// Supports: "=X", "<>X", ">N", ">=N", "<N", "<=N", "*wild*", plain text.
bool FunctionRegistry::matchCriteria(const QVariant& val, const QString& criteria) {
    QString c = criteria.trimmed();
    auto testNum = [&](auto cmp) -> bool {
        bool ok1, ok2;
        double lhs = val.toDouble(&ok1);
        double rhs = c.mid(c.startsWith("<>") || c.startsWith(">=") || c.startsWith("<=") ? 2 : 1).toDouble(&ok2);
        return ok1 && ok2 && cmp(lhs, rhs);
    };
    if (c.startsWith(">=")) return testNum([](double a, double b){ return a >= b; });
    if (c.startsWith("<=")) return testNum([](double a, double b){ return a <= b; });
    if (c.startsWith("<>")) return val.toString().compare(c.mid(2), Qt::CaseInsensitive) != 0;
    if (c.startsWith(">"))  return testNum([](double a, double b){ return a >  b; });
    if (c.startsWith("<"))  return testNum([](double a, double b){ return a <  b; });
    if (c.startsWith("="))  return val.toString().compare(c.mid(1), Qt::CaseInsensitive) == 0;
    if (c.contains('*') || c.contains('?')) {
        QRegularExpression re(QRegularExpression::wildcardToRegularExpression(c),
                              QRegularExpression::CaseInsensitiveOption);
        return re.match(val.toString()).hasMatch();
    }
    return val.toString().compare(c, Qt::CaseInsensitive) == 0;
}

bool FunctionRegistry::toBool(const QVariant& v) {
    if (v.typeId()==QMetaType::Bool) return v.toBool();
    if (v.canConvert<double>()) return v.toDouble() != 0.0;
    QString s = v.toString();
    return !s.isEmpty() && s != "0" && s.toLower() != "false";
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
    registerFn("COUNTA", [this](QList<QVariant> a) -> QVariant {
        int cnt = 0;
        for (const auto& v : a) {
            auto flat = flattenToList(v);
            for (const auto& item : flat)
                if (!item.isNull() && item.isValid() && !item.toString().isEmpty()) cnt++;
        }
        return (double)cnt;
    });
    registerFn("COUNTBLANK", [this](QList<QVariant> a) -> QVariant {
        int cnt = 0;
        for (const auto& v : a) {
            auto flat = flattenToList(v);
            for (const auto& item : flat)
                if (item.isNull() || !item.isValid() || item.toString().isEmpty()) cnt++;
        }
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
        // VLOOKUP(lookup, table_2d, col_index, [exact_match=TRUE])
        if (a.size() < 3) return QString("#N/A");
        QVariant lookup = a.value(0);
        QList<QVariant> rows = a.value(1).toList(); // list-of-rows from expandRange2D
        int colIdx = (int)toNum(a.value(2)) - 1;
        bool exact = a.size() < 4 || toBool(a.value(3));
        for (const auto& rowVar : rows) {
            QList<QVariant> row = rowVar.toList();
            if (row.isEmpty()) continue;
            bool match = exact
                ? (row[0].toString().compare(lookup.toString(), Qt::CaseInsensitive) == 0 ||
                   (row[0].canConvert<double>() && lookup.canConvert<double>() &&
                    toNum(row[0]) == toNum(lookup)))
                : (row[0].toString().compare(lookup.toString(), Qt::CaseInsensitive) == 0 ||
                   (row[0].canConvert<double>() && lookup.canConvert<double>() &&
                    toNum(row[0]) == toNum(lookup)));
            if (match) {
                if (colIdx >= 0 && colIdx < row.size()) return row[colIdx];
                return QString("#REF!");
            }
        }
        return QString("#N/A");
    });
    registerFn("HLOOKUP", [this](QList<QVariant> a) -> QVariant {
        // HLOOKUP(lookup, table_2d, row_index, [exact_match=TRUE])
        if (a.size() < 3) return QString("#N/A");
        QVariant lookup = a.value(0);
        QList<QVariant> rows = a.value(1).toList();
        int rowIdx = (int)toNum(a.value(2)) - 1;
        if (rows.isEmpty()) return QString("#N/A");
        // Search first row (header row) for lookup value
        QList<QVariant> headerRow = rows[0].toList();
        int foundCol = -1;
        for (int c = 0; c < headerRow.size(); ++c) {
            bool match = headerRow[c].toString().compare(lookup.toString(), Qt::CaseInsensitive) == 0 ||
                         (headerRow[c].canConvert<double>() && lookup.canConvert<double>() &&
                          toNum(headerRow[c]) == toNum(lookup));
            if (match) { foundCol = c; break; }
        }
        if (foundCol < 0) return QString("#N/A");
        if (rowIdx < 0 || rowIdx >= rows.size()) return QString("#REF!");
        QList<QVariant> targetRow = rows[rowIdx].toList();
        if (foundCol >= targetRow.size()) return QString("#REF!");
        return targetRow[foundCol];
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

// ── Additional functions appended to registerAll() ───────────────────────────
// NOTE: These are added via a separate registerExtended() call to avoid
// modifying the original registerAll() signature.
void FunctionRegistry::registerExtended() {
    // ── Statistical ──────────────────────────────────────────────────────────
    registerFn("MEDIAN", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.isEmpty()) return 0.0;
        std::sort(nums.begin(), nums.end());
        int n = nums.size();
        return (n%2==0) ? (nums[n/2-1]+nums[n/2])/2.0 : nums[n/2];
    });
    registerFn("STDEV", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.size() < 2) return QString("#NUM!");
        double mean = std::accumulate(nums.begin(),nums.end(),0.0)/nums.size();
        double var = 0; for (double x : nums) var += (x-mean)*(x-mean);
        return std::sqrt(var/(nums.size()-1));
    });
    registerFn("VAR", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        if (nums.size() < 2) return QString("#NUM!");
        double mean = std::accumulate(nums.begin(),nums.end(),0.0)/nums.size();
        double var = 0; for (double x : nums) var += (x-mean)*(x-mean);
        return var/(nums.size()-1);
    });
    registerFn("LARGE", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        int k = (int)toNum(a.value(a.size()-1));
        if (k<1||k>(int)nums.size()) return QString("#NUM!");
        std::sort(nums.rbegin(),nums.rend());
        return nums[k-1];
    });
    registerFn("SMALL", [this](QList<QVariant> a) -> QVariant {
        auto nums = toNums(a);
        int k = (int)toNum(a.value(a.size()-1));
        if (k<1||k>(int)nums.size()) return QString("#NUM!");
        std::sort(nums.begin(),nums.end());
        return nums[k-1];
    });
    registerFn("RANK", [this](QList<QVariant> a) -> QVariant {
        double val = toNum(a.value(0));
        auto nums = toNums(a.mid(1,a.size()-1));
        bool asc = a.size()>2 && toNum(a.last())!=0;
        if (asc) std::sort(nums.begin(),nums.end());
        else     std::sort(nums.rbegin(),nums.rend());
        auto it = std::find(nums.begin(),nums.end(),val);
        if (it==nums.end()) return QString("#N/A");
        return (double)(std::distance(nums.begin(),it)+1);
    });

    // ── Lookup / Reference ───────────────────────────────────────────────────
    // VLOOKUP/HLOOKUP fully implemented in registerAll(); these are no-ops to avoid double-registration.
    // (registerAll runs before registerExtended so the correct versions are already registered.)
    registerFn("INDEX", [](QList<QVariant> a) -> QVariant {
        int idx = (int)a.value(1).toDouble() - 1;
        if (idx < 0 || idx >= a.size()-1) return QString("#REF!");
        return a.value(idx);
    });
    registerFn("MATCH", [this](QList<QVariant> a) -> QVariant {
        if (a.size() < 2) return QString("#VALUE!");
        QVariant needle = a.value(0);
        for (int i=1; i<a.size(); ++i)
            if (a[i].toString().compare(needle.toString(),Qt::CaseInsensitive)==0)
                return (double)i;
        return QString("#N/A");
    });
    registerFn("CHOOSE", [](QList<QVariant> a) -> QVariant {
        int idx = (int)a.value(0).toDouble();
        if (idx<1||idx>=a.size()) return QString("#VALUE!");
        return a.value(idx);
    });

    // ── Text functions ────────────────────────────────────────────────────────
    registerFn("CONCATENATE", [](QList<QVariant> a) -> QVariant {
        QString s; for (auto& v:a) s+=v.toString(); return s;
    });
    registerFn("CONCAT", [](QList<QVariant> a) -> QVariant {
        QString s; for (auto& v:a) s+=v.toString(); return s;
    });
    registerFn("LEFT", [this](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().left((int)toNum(a.value(1,1)));
    });
    registerFn("RIGHT", [this](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().right((int)toNum(a.value(1,1)));
    });
    registerFn("MID", [this](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().mid((int)toNum(a.value(1))-1, (int)toNum(a.value(2)));
    });
    registerFn("LEN", [](QList<QVariant> a) -> QVariant {
        return (double)a.value(0).toString().length();
    });
    registerFn("UPPER", [](QList<QVariant> a) -> QVariant { return a.value(0).toString().toUpper(); });
    registerFn("LOWER", [](QList<QVariant> a) -> QVariant { return a.value(0).toString().toLower(); });
    registerFn("TRIM",  [](QList<QVariant> a) -> QVariant { return a.value(0).toString().trimmed(); });
    registerFn("SUBSTITUTE",[](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().replace(a.value(1).toString(),a.value(2).toString());
    });
    registerFn("FIND", [](QList<QVariant> a) -> QVariant {
        int pos = a.value(1).toString().indexOf(a.value(0).toString());
        return pos<0 ? QVariant(QString("#VALUE!")) : QVariant((double)pos+1);
    });
    registerFn("TEXT", [this](QList<QVariant> a) -> QVariant {
        double v = toNum(a.value(0));
        QString fmt = a.value(1).toString();
        // Basic format codes
        if (fmt=="0") return QString::number((long long)v);
        if (fmt.startsWith("0.")) {
            int dec = fmt.length()-2;
            return QString::number(v,'f',dec);
        }
        return QString::number(v);
    });
    registerFn("VALUE", [this](QList<QVariant> a) -> QVariant {
        bool ok; double d=a.value(0).toString().toDouble(&ok);
        return ok ? QVariant(d) : QVariant(QString("#VALUE!"));
    });
    registerFn("REPT", [this](QList<QVariant> a) -> QVariant {
        return a.value(0).toString().repeated((int)toNum(a.value(1)));
    });

    // ── Logical functions ─────────────────────────────────────────────────────
    registerFn("AND", [this](QList<QVariant> a) -> QVariant {
        for (auto& v:a) if (!toBool(v)) return false; return true;
    });
    registerFn("OR", [this](QList<QVariant> a) -> QVariant {
        for (auto& v:a) if (toBool(v)) return true; return false;
    });
    registerFn("NOT", [this](QList<QVariant> a) -> QVariant { return !toBool(a.value(0)); });
    registerFn("IFERROR", [](QList<QVariant> a) -> QVariant {
        QString s = a.value(0).toString();
        if (s.startsWith('#')) return a.value(1);
        return a.value(0);
    });
    registerFn("IFNA", [](QList<QVariant> a) -> QVariant {
        return a.value(0).toString()=="#N/A" ? a.value(1) : a.value(0);
    });
    registerFn("IFS", [this](QList<QVariant> a) -> QVariant {
        for (int i=0; i+1<a.size(); i+=2)
            if (toBool(a[i])) return a[i+1];
        return QString("#N/A");
    });
    registerFn("SWITCH", [this](QList<QVariant> a) -> QVariant {
        QVariant expr = a.value(0);
        for (int i=1; i+1<a.size(); i+=2)
            if (expr==a[i]) return a[i+1];
        if (a.size()%2==0) return a.last(); // default
        return QString("#N/A");
    });
    registerFn("TRUE",  [](QList<QVariant>) -> QVariant { return true; });
    registerFn("FALSE", [](QList<QVariant>) -> QVariant { return false; });
    registerFn("NA",    [](QList<QVariant>) -> QVariant { return QString("#N/A"); });

    // ── Math ─────────────────────────────────────────────────────────────────
    registerFn("SUMIF", [this](QList<QVariant> a) -> QVariant {
        // SUMIF(range, criteria, [sum_range])
        if (a.size() < 2) return 0.0;
        QList<QVariant> rangeVals = flattenToList(a.value(0));
        QString criteria = a.value(1).toString();
        QList<QVariant> sumVals = a.size() > 2 ? flattenToList(a.value(2)) : rangeVals;
        double sum = 0.0;
        for (int i = 0; i < rangeVals.size(); ++i) {
            if (matchCriteria(rangeVals[i], criteria) && i < sumVals.size())
                sum += toNum(sumVals[i]);
        }
        return sum;
    });
    registerFn("SUMIFS", [this](QList<QVariant> a) -> QVariant {
        // SUMIFS(sum_range, range1, crit1, range2, crit2, ...)
        if (a.size() < 3) return 0.0;
        QList<QVariant> sumVals = flattenToList(a.value(0));
        double sum = 0.0;
        for (int i = 0; i < sumVals.size(); ++i) {
            bool allMatch = true;
            for (int p = 1; p + 1 < a.size(); p += 2) {
                QList<QVariant> rangeVals = flattenToList(a.value(p));
                QString crit = a.value(p + 1).toString();
                if (i >= rangeVals.size() || !matchCriteria(rangeVals[i], crit)) {
                    allMatch = false;
                    break;
                }
            }
            if (allMatch) sum += toNum(sumVals[i]);
        }
        return sum;
    });
    registerFn("COUNTIF", [this](QList<QVariant> a) -> QVariant {
        // COUNTIF(range, criteria)
        if (a.size() < 2) return 0.0;
        QList<QVariant> rangeVals = flattenToList(a.value(0));
        QString criteria = a.value(1).toString();
        int cnt = 0;
        for (const auto& v : rangeVals)
            if (matchCriteria(v, criteria)) cnt++;
        return (double)cnt;
    });
    registerFn("COUNTIFS", [this](QList<QVariant> a) -> QVariant {
        // COUNTIFS(range1, crit1, range2, crit2, ...)
        if (a.size() < 2) return 0.0;
        QList<QVariant> first = flattenToList(a.value(0));
        int cnt = 0;
        for (int i = 0; i < first.size(); ++i) {
            bool allMatch = true;
            for (int p = 0; p + 1 < a.size(); p += 2) {
                QList<QVariant> rangeVals = flattenToList(a.value(p));
                QString crit = a.value(p + 1).toString();
                if (i >= rangeVals.size() || !matchCriteria(rangeVals[i], crit)) {
                    allMatch = false;
                    break;
                }
            }
            if (allMatch) cnt++;
        }
        return (double)cnt;
    });
    registerFn("AVERAGEIF", [this](QList<QVariant> a) -> QVariant {
        // AVERAGEIF(range, criteria, [avg_range])
        if (a.size() < 2) return QString("#DIV/0!");
        QList<QVariant> rangeVals = flattenToList(a.value(0));
        QString criteria = a.value(1).toString();
        QList<QVariant> avgVals = a.size() > 2 ? flattenToList(a.value(2)) : rangeVals;
        double sum = 0.0; int cnt = 0;
        for (int i = 0; i < rangeVals.size(); ++i) {
            if (matchCriteria(rangeVals[i], criteria) && i < avgVals.size()) {
                sum += toNum(avgVals[i]); cnt++;
            }
        }
        return cnt > 0 ? sum / cnt : QVariant(QString("#DIV/0!"));
    });
    registerFn("AVERAGEIFS", [this](QList<QVariant> a) -> QVariant {
        // AVERAGEIFS(avg_range, range1, crit1, range2, crit2, ...)
        if (a.size() < 3) return QString("#DIV/0!");
        QList<QVariant> avgVals = flattenToList(a.value(0));
        double sum = 0.0; int cnt = 0;
        for (int i = 0; i < avgVals.size(); ++i) {
            bool allMatch = true;
            for (int p = 1; p + 1 < a.size(); p += 2) {
                QList<QVariant> rangeVals = flattenToList(a.value(p));
                QString crit = a.value(p + 1).toString();
                if (i >= rangeVals.size() || !matchCriteria(rangeVals[i], crit)) {
                    allMatch = false;
                    break;
                }
            }
            if (allMatch) { sum += toNum(avgVals[i]); cnt++; }
        }
        return cnt > 0 ? sum / cnt : QVariant(QString("#DIV/0!"));
    });
    registerFn("MOD",  [this](QList<QVariant> a) -> QVariant {
        double b=toNum(a.value(1)); if(b==0) return QString("#DIV/0!");
        return std::fmod(toNum(a.value(0)),b);
    });
    registerFn("INT",  [this](QList<QVariant> a) -> QVariant { return std::floor(toNum(a.value(0))); });
    registerFn("TRUNC",[this](QList<QVariant> a) -> QVariant { return std::trunc(toNum(a.value(0))); });
    registerFn("CEILING",[this](QList<QVariant> a) -> QVariant {
        double v=toNum(a.value(0)), sig=toNum(a.value(1,1));
        return std::ceil(v/sig)*sig;
    });
    registerFn("FLOOR",[this](QList<QVariant> a) -> QVariant {
        double v=toNum(a.value(0)), sig=toNum(a.value(1,1));
        return std::floor(v/sig)*sig;
    });
    registerFn("LN",   [this](QList<QVariant> a) -> QVariant {
        double v=toNum(a.value(0)); if(v<=0) return QString("#NUM!"); return std::log(v);
    });
    registerFn("LOG",  [this](QList<QVariant> a) -> QVariant {
        double v=toNum(a.value(0)), base=toNum(a.value(1,10));
        if(v<=0||base<=0) return QString("#NUM!"); return std::log(v)/std::log(base);
    });
    registerFn("LOG10",[this](QList<QVariant> a) -> QVariant {
        double v=toNum(a.value(0)); if(v<=0) return QString("#NUM!"); return std::log10(v);
    });
    registerFn("EXP",  [this](QList<QVariant> a) -> QVariant { return std::exp(toNum(a.value(0))); });
    registerFn("SIN",  [this](QList<QVariant> a) -> QVariant { return std::sin(toNum(a.value(0))); });
    registerFn("COS",  [this](QList<QVariant> a) -> QVariant { return std::cos(toNum(a.value(0))); });
    registerFn("TAN",  [this](QList<QVariant> a) -> QVariant { return std::tan(toNum(a.value(0))); });
    registerFn("ASIN", [this](QList<QVariant> a) -> QVariant { return std::asin(toNum(a.value(0))); });
    registerFn("ACOS", [this](QList<QVariant> a) -> QVariant { return std::acos(toNum(a.value(0))); });
    registerFn("ATAN", [this](QList<QVariant> a) -> QVariant { return std::atan(toNum(a.value(0))); });
    registerFn("ATAN2",[this](QList<QVariant> a) -> QVariant {
        return std::atan2(toNum(a.value(0)),toNum(a.value(1)));
    });
    registerFn("PI",   [](QList<QVariant>) -> QVariant { return M_PI; });
    registerFn("E",    [](QList<QVariant>) -> QVariant { return M_E; });
    registerFn("RAND", [](QList<QVariant>) -> QVariant {
        return (double)rand()/(double)RAND_MAX;
    });
    registerFn("RANDBETWEEN",[this](QList<QVariant> a) -> QVariant {
        int lo=(int)toNum(a.value(0)), hi=(int)toNum(a.value(1));
        if(lo>hi) return QString("#NUM!");
        return (double)(lo + rand()%(hi-lo+1));
    });
    registerFn("PRODUCT",[this](QList<QVariant> a) -> QVariant {
        double p=1; for(double n:toNums(a)) p*=n; return p;
    });
    registerFn("SUMPRODUCT",[this](QList<QVariant> a) -> QVariant {
        // Flat arg version: sum of products of consecutive pairs
        auto nums=toNums(a); double s=0;
        for(int i=0;i+1<nums.size();i+=2) s+=nums[i]*nums[i+1];
        return s;
    });
    registerFn("GCD",[this](QList<QVariant> a) -> QVariant {
        auto nums=toNums(a);
        if(nums.empty()) return 0.0;
        long long g=(long long)nums[0];
        for(int i=1;i<nums.size();i++){
            long long b=(long long)nums[i];
            while(b){long long t=b;b=g%b;g=t;}
        }
        return (double)g;
    });
    registerFn("LCM",[this](QList<QVariant> a) -> QVariant {
        auto nums=toNums(a);
        if(nums.empty()) return 0.0;
        long long l=(long long)std::abs(nums[0]);
        for(int i=1;i<nums.size();i++){
            long long b=(long long)std::abs(nums[i]);
            long long g=l; long long r=b;
            while(r){long long t=r;r=g%r;g=t;}
            l=l/g*b;
        }
        return (double)l;
    });
    registerFn("FACT",[this](QList<QVariant> a) -> QVariant {
        int n=(int)toNum(a.value(0)); if(n<0) return QString("#NUM!");
        double f=1; for(int i=2;i<=n;i++) f*=i; return f;
    });

    // ── Date functions ────────────────────────────────────────────────────────
    registerFn("DATE", [this](QList<QVariant> a) -> QVariant {
        return QDate((int)toNum(a.value(0)),(int)toNum(a.value(1)),(int)toNum(a.value(2))).toString(Qt::ISODate);
    });
    registerFn("DATEDIF",[this](QList<QVariant> a) -> QVariant {
        QDate d1=QDate::fromString(a.value(0).toString(),Qt::ISODate);
        QDate d2=QDate::fromString(a.value(1).toString(),Qt::ISODate);
        QString unit=a.value(2).toString().toUpper();
        if(unit=="D") return (double)d1.daysTo(d2);
        if(unit=="M") return (double)(d1.month()-d2.month()+(d1.year()-d2.year())*12);
        if(unit=="Y") return (double)(d1.year()-d2.year());
        return QString("#VALUE!");
    });
    registerFn("EDATE",[this](QList<QVariant> a) -> QVariant {
        QDate d=QDate::fromString(a.value(0).toString(),Qt::ISODate);
        return d.addMonths((int)toNum(a.value(1))).toString(Qt::ISODate);
    });
    registerFn("WEEKDAY",[this](QList<QVariant> a) -> QVariant {
        QDate d=QDate::fromString(a.value(0).toString(),Qt::ISODate);
        return (double)d.dayOfWeek();
    });
    registerFn("WEEKNUM",[this](QList<QVariant> a) -> QVariant {
        QDate d=QDate::fromString(a.value(0).toString(),Qt::ISODate);
        return (double)d.weekNumber();
    });
    registerFn("DAYS",[this](QList<QVariant> a) -> QVariant {
        QDate d2=QDate::fromString(a.value(0).toString(),Qt::ISODate);
        QDate d1=QDate::fromString(a.value(1).toString(),Qt::ISODate);
        return (double)d1.daysTo(d2);
    });
    registerFn("NETWORKDAYS",[this](QList<QVariant> a) -> QVariant {
        QDate d1=QDate::fromString(a.value(0).toString(),Qt::ISODate);
        QDate d2=QDate::fromString(a.value(1).toString(),Qt::ISODate);
        int days=0, step=d1<=d2?1:-1;
        for(QDate d=d1; d!=d2; d=d.addDays(step))
            if(d.dayOfWeek()<=5) days+=step;
        return (double)days;
    });

    // ── Financial ────────────────────────────────────────────────────────────
    registerFn("PMT",[this](QList<QVariant> a) -> QVariant {
        double rate=toNum(a.value(0)), nper=toNum(a.value(1)), pv=toNum(a.value(2));
        if(rate==0) return -pv/nper;
        double pmt = rate*pv/(1-std::pow(1+rate,-nper));
        return -pmt;
    });
    registerFn("PV",[this](QList<QVariant> a) -> QVariant {
        double rate=toNum(a.value(0)), nper=toNum(a.value(1)), pmt=toNum(a.value(2));
        if(rate==0) return -pmt*nper;
        return -pmt*(1-std::pow(1+rate,-nper))/rate;
    });
    registerFn("FV",[this](QList<QVariant> a) -> QVariant {
        double rate=toNum(a.value(0)), nper=toNum(a.value(1)), pmt=toNum(a.value(2));
        double pv=toNum(a.value(3,0));
        return -pmt*(std::pow(1+rate,nper)-1)/rate - pv*std::pow(1+rate,nper);
    });
    registerFn("NPV",[this](QList<QVariant> a) -> QVariant {
        double rate=toNum(a.value(0)), npv=0;
        for(int i=1;i<a.size();i++) npv+=toNum(a[i])/std::pow(1+rate,i);
        return npv;
    });
}
