#pragma once
// ─────────────────────────────────────────────────────────────────────────────
//  PluginManager — discovers, loads, and manages OpenSheet plugins.
//
//  Usage:
//      PluginManager mgr;
//      mgr.setCore(core);
//      mgr.loadPluginsFromDirectory("plugins/");
//      // All custom formulas from loaded plugins are already registered.
//
//  Plugin search order inside the directory:
//      *.dll   (Windows C++ plugins)
//      *.so    (Linux / macOS C++ plugins)
//      *.dylib (macOS C++ plugins)
// ─────────────────────────────────────────────────────────────────────────────
#ifdef _WIN32
  #ifdef PLUGINSYSTEM_EXPORTS
    #define PSYS_API __declspec(dllexport)
  #else
    #define PSYS_API __declspec(dllimport)
  #endif
#else
  #define PSYS_API __attribute__((visibility("default")))
#endif

#include <QObject>
#include <QString>
#include <QStringList>
#include <QList>
#include <memory>

class IPlugin;
class ISpreadsheetCore;

// Forward-declare — we own QLibrary objects through void* to avoid exposing
// the Qt header here.
struct PluginEntry;

class PSYS_API PluginManager : public QObject
{
    Q_OBJECT
public:
    explicit PluginManager(QObject* parent = nullptr);
    ~PluginManager() override;

    // Set the spreadsheet core passed to each plugin's initialize().
    void setCore(ISpreadsheetCore* core);

    // Load a single plugin from an absolute path.  Returns true on success.
    bool loadPlugin(const QString& filePath);

    // Scan a directory and load every file that looks like a plugin.
    int loadPluginsFromDirectory(const QString& dirPath);

    // Unload and destroy a plugin by name.  Returns true if found and removed.
    bool unloadPlugin(const QString& name);

    // Unload all plugins (called automatically on destruction).
    void unloadAll();

    // Access loaded plugins.
    QList<IPlugin*> plugins() const;
    IPlugin*        pluginByName(const QString& name) const;

    // List of directories that have been scanned.
    QStringList pluginDirectories() const;

    // Last error message from loadPlugin().
    QString lastError() const;

signals:
    void pluginLoaded(const QString& name);
    void pluginUnloaded(const QString& name);
    void pluginError(const QString& filePath, const QString& error);

private:
    struct Impl;
    Impl* d;
};

extern "C" PSYS_API PluginManager* createPluginManager();
