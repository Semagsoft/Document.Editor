#include "PluginManager.h"
#include "IPlugin.h"

#include <QCoreApplication>
#include <QDir>
#include <QFileInfo>
#include <QLibrary>

PluginManager::PluginManager(QObject* parent)
    : QObject(parent)
{
}

void PluginManager::loadPlugins(const QString& directory)
{
    m_directory = directory.isEmpty()
        ? QDir(QCoreApplication::applicationDirPath() + QStringLiteral("/plugins")).absolutePath()
        : directory;

    QDir pluginsDir(m_directory);
    if (!pluginsDir.exists()) {
        qDebug() << "Plugin directory does not exist:" << m_directory;
        return;
    }

    const QStringList entries = pluginsDir.entryList(QDir::Files | QDir::NoDotAndDotDot);
    qDebug() << "Found" << entries.size() << "plugin files in" << m_directory;
    
    for (const QString& entry : entries) {
        QString filePath = pluginsDir.absoluteFilePath(entry);
        if (QLibrary::isLibrary(filePath))
            loadPlugin(filePath);
        else
            qWarning() << "Skipping non-plugin file:" << entry;
    }
}

void PluginManager::loadPlugin(const QString& filePath)
{
    qDebug() << "Loading plugin from:" << filePath;
    
    auto* loader = new QPluginLoader(filePath);
    QObject* instance = loader->instance();
    if (!instance) {
        QString pluginName = QFileInfo(filePath).baseName();
        qWarning() << "Failed to load plugin" << pluginName << ":" << loader->errorString();
        emit pluginError(pluginName, loader->errorString());
        delete loader;
        return;
    }

    IPlugin* plugin = qobject_cast<IPlugin*>(instance);
    if (!plugin) {
        QString pluginName = QFileInfo(filePath).baseName();
        qWarning() << "Loaded object is not a valid plugin:" << pluginName;
        delete loader;
        emit pluginError(pluginName, tr("Not a valid plugin"));
        return;
    }

    qDebug() << "Plugin metadata - Name:" << plugin->name() << "Version:" << plugin->version() 
             << "Description:" << plugin->description();
    
    if (!plugin->initialize(m_context)) {
        qWarning() << "Plugin initialization failed:" << plugin->name();
        delete loader;
        emit pluginError(plugin->name(), tr("Plugin failed to initialize"));
        return;
    }

    PluginInfo info;
    info.name = plugin->name();
    info.version = plugin->version();
    info.description = plugin->description();
    info.filePath = filePath;
    info.instance = plugin;
    info.loader = loader;
    info.loaded = true;

    m_plugins.append(info);
    qDebug() << "Successfully loaded plugin:" << info.name;
    emit pluginLoaded(info.name);
}

PluginContext* PluginManager::context() const
{
    return m_context;
}

void PluginManager::setContext(PluginContext* context)
{
    m_context = context;
}

void PluginManager::unloadPlugins()
{
    for (auto it = m_plugins.begin(); it != m_plugins.end(); ++it) {
        if (it->instance)
            it->instance->shutdown();
        emit pluginUnloaded(it->name);
        if (it->loader) {
            it->loader->unload();
            delete it->loader;
        }
    }
    m_plugins.clear();
}

QList<PluginManager::PluginInfo> PluginManager::plugins() const
{
    return m_plugins;
}

int PluginManager::pluginCount() const
{
    return m_plugins.size();
}

bool PluginManager::hasPlugin(const QString& name) const
{
    for (const auto& p : m_plugins) {
        if (p.name == name)
            return true;
    }
    return false;
}

QStringList PluginManager::pluginNames() const
{
    QStringList names;
    for (const auto& p : m_plugins)
        names.append(p.name);
    return names;
}

QStringList PluginManager::loadedPluginNames() const
{
    QStringList names;
    for (const auto& p : m_plugins) {
        if (p.loaded)
            names.append(p.name);
    }
    return names;
}
