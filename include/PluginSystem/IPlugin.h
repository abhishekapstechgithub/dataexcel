#pragma once
// ─────────────────────────────────────────────────────────────────────────────
//  IPlugin — base interface every OpenSheet plugin must implement.
//
//  C++ plugins are compiled as shared libraries (.dll / .so) that export a
//  factory function:
//
//      extern "C" PLUGIN_API IPlugin* createPlugin();
//
//  Python plugins are loaded via a thin C++ wrapper that calls the Python
//  interpreter.  See PluginManager for loading details.
// ─────────────────────────────────────────────────────────────────────────────
#ifdef _WIN32
  #ifdef PLUGIN_EXPORTS
    #define PLUGIN_API __declspec(dllexport)
  #else
    #define PLUGIN_API __declspec(dllimport)
  #endif
#else
  #define PLUGIN_API __attribute__((visibility("default")))
#endif

#include <QString>
#include <QVariant>
#include <QMap>
#include <functional>

class ISpreadsheetCore;

// Callable type for custom formula functions registered by plugins.
using PluginFnCallable = std::function<QVariant(QList<QVariant>)>;

class IPlugin
{
public:
    virtual ~IPlugin() = default;

    // ── Identity ──────────────────────────────────────────────────────────────
    virtual QString name()        const = 0;
    virtual QString version()     const = 0;
    virtual QString description() const = 0;

    // ── Life-cycle ────────────────────────────────────────────────────────────
    // Called once after the plugin is loaded.  core may be nullptr during
    // formula-only plugins that don't need data access.
    virtual void initialize(ISpreadsheetCore* core) = 0;

    // Called before the plugin is unloaded.
    virtual void shutdown() = 0;

    // ── Custom formulas (optional) ────────────────────────────────────────────
    // Return a map of UPPERCASE function name → callable.
    // The PluginManager registers these with the FormulaEngine automatically.
    virtual QMap<QString, PluginFnCallable> customFunctions() const { return {}; }
};

// ── Factory function that every plugin DLL must export ────────────────────────
extern "C" PLUGIN_API IPlugin* createPlugin();
