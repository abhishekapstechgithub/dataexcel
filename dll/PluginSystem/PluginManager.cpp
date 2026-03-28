#include "PluginManager.h"
#include "IPlugin.h"
#include <QLibrary>
#include <QDir>
#include <QFileInfo>

// ─────────────────────────────────────────────────────────────────────────────
//  Impl
// ─────────────────────────────────────────────────────────────────────────────
struct PluginManager::Impl {
    ISpreadsheetCore*     core { nullptr };
    QString               lastError;
    QStringList           pluginDirs;

    struct Entry {
        QLibrary* lib    { nullptr };
        IPlugin*  plugin { nullptr };
    };
    QList<Entry> entries;
};

// ─────────────────────────────────────────────────────────────────────────────
PluginManager::PluginManager(QObject* parent)
    : QObject(parent), d(new Impl) {}

PluginManager::~PluginManager()
{
    unloadAll();
    delete d;
}

void PluginManager::setCore(ISpreadsheetCore* core)
{
    d->core = core;
}

// ─────────────────────────────────────────────────────────────────────────────
bool PluginManager::loadPlugin(const QString& filePath)
{
    QFileInfo fi(filePath);
    if (!fi.exists()) {
        d->lastError = "File not found: " + filePath;
        emit pluginError(filePath, d->lastError);
        return false;
    }

    auto* lib = new QLibrary(filePath, this);
    if (!lib->load()) {
        d->lastError = lib->errorString();
        emit pluginError(filePath, d->lastError);
        delete lib;
        return false;
    }

    using FactoryFn = IPlugin* (*)();
    auto factory = reinterpret_cast<FactoryFn>(lib->resolve("createPlugin"));
    if (!factory) {
        d->lastError = "Symbol 'createPlugin' not found in: " + filePath;
        emit pluginError(filePath, d->lastError);
        lib->unload();
        delete lib;
        return false;
    }

    IPlugin* plugin = factory();
    if (!plugin) {
        d->lastError = "createPlugin() returned nullptr for: " + filePath;
        emit pluginError(filePath, d->lastError);
        lib->unload();
        delete lib;
        return false;
    }

    plugin->initialize(d->core);

    d->entries.append({ lib, plugin });
    emit pluginLoaded(plugin->name());
    return true;
}

int PluginManager::loadPluginsFromDirectory(const QString& dirPath)
{
    d->pluginDirs.append(dirPath);
    QDir dir(dirPath);
    if (!dir.exists()) return 0;

#ifdef Q_OS_WIN
    QStringList filters = { "*.dll" };
#elif defined(Q_OS_MAC)
    QStringList filters = { "*.dylib", "*.so" };
#else
    QStringList filters = { "*.so" };
#endif

    int loaded = 0;
    for (const QFileInfo& fi : dir.entryInfoList(filters, QDir::Files)) {
        if (loadPlugin(fi.absoluteFilePath()))
            ++loaded;
    }
    return loaded;
}

bool PluginManager::unloadPlugin(const QString& name)
{
    for (int i = 0; i < d->entries.size(); ++i) {
        auto& e = d->entries[i];
        if (e.plugin && e.plugin->name() == name) {
            e.plugin->shutdown();
            delete e.plugin;
            e.lib->unload();
            delete e.lib;
            d->entries.removeAt(i);
            emit pluginUnloaded(name);
            return true;
        }
    }
    return false;
}

void PluginManager::unloadAll()
{
    for (auto& e : d->entries) {
        if (e.plugin) { e.plugin->shutdown(); delete e.plugin; }
        if (e.lib)    { e.lib->unload(); delete e.lib; }
    }
    d->entries.clear();
}

QList<IPlugin*> PluginManager::plugins() const
{
    QList<IPlugin*> out;
    for (const auto& e : d->entries)
        if (e.plugin) out << e.plugin;
    return out;
}

IPlugin* PluginManager::pluginByName(const QString& name) const
{
    for (const auto& e : d->entries)
        if (e.plugin && e.plugin->name() == name) return e.plugin;
    return nullptr;
}

QStringList PluginManager::pluginDirectories() const { return d->pluginDirs; }
QString     PluginManager::lastError()          const { return d->lastError; }

#include "PluginManager.moc"

extern "C" PSYS_API PluginManager* createPluginManager() {
    return new PluginManager();
}
