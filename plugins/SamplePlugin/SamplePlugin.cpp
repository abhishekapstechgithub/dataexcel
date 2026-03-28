// ─────────────────────────────────────────────────────────────────────────────
//  SamplePlugin — demonstrates the OpenSheet plugin API.
//
//  Registers two custom spreadsheet functions:
//    HELLO()        → returns "Hello from SamplePlugin!"
//    DOUBLE(n)      → returns n * 2
// ─────────────────────────────────────────────────────────────────────────────
#include "IPlugin.h"
#include <QDebug>

class SamplePlugin : public IPlugin
{
public:
    QString name()        const override { return "SamplePlugin"; }
    QString version()     const override { return "1.0.0"; }
    QString description() const override { return "Demonstration plugin with custom formulas"; }

    void initialize(ISpreadsheetCore* /*core*/) override {
        qDebug() << "[SamplePlugin] initialized";
    }

    void shutdown() override {
        qDebug() << "[SamplePlugin] shutdown";
    }

    QMap<QString, PluginFnCallable> customFunctions() const override {
        QMap<QString, PluginFnCallable> fns;

        fns["HELLO"] = [](QList<QVariant>) -> QVariant {
            return QString("Hello from SamplePlugin!");
        };

        fns["DOUBLE"] = [](QList<QVariant> args) -> QVariant {
            bool ok;
            double n = args.value(0).toDouble(&ok);
            if (!ok) return QString("#VALUE!");
            return n * 2.0;
        };

        return fns;
    }
};

extern "C" PLUGIN_API IPlugin* createPlugin() {
    return new SamplePlugin();
}
