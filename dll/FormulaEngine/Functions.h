#pragma once
#include <QVariant>
#include <QHash>
#include <QString>
#include <functional>

using FnCallable = std::function<QVariant(QList<QVariant>)>;

class FunctionRegistry {
public:
    void registerAll();
    void registerFn(const QString& name, FnCallable fn);
    QVariant call(const QString& name, const QList<QVariant>& args);
    bool has(const QString& name) const;
private:
    QHash<QString, FnCallable> m_fns;
    static double toNum(const QVariant& v);
    static QList<double> toNums(const QList<QVariant>& args);
};
