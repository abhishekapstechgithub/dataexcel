#pragma once
#include <QVariant>
#include <QHash>
#include <QString>
#include <functional>

using FnCallable = std::function<QVariant(QList<QVariant>)>;

class FunctionRegistry {
public:
    void registerAll();
    void registerExtended();  // 60+ additional functions
    void registerFn(const QString& name, FnCallable fn);
    QVariant call(const QString& name, const QList<QVariant>& args);
    bool has(const QString& name) const;
private:
    QHash<QString, FnCallable> m_fns;
    static bool   toBool(const QVariant& v);
    static double toNum(const QVariant& v);
    static QList<double> toNums(const QList<QVariant>& args);
};
