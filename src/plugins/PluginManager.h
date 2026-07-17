#pragma once

#include <QList>
#include <QObject>
#include <QPluginLoader>
#include <QString>
#include <QStringList>

class IPlugin;
class PluginContext;

class PluginManager : public QObject {
    Q_OBJECT

public:
    struct PluginInfo {
        QString name;
        QString version;
        QString description;
        QString filePath;
        IPlugin* instance = nullptr;
        QPluginLoader* loader = nullptr;
        bool loaded = false;
    };

    explicit PluginManager(QObject* parent = nullptr);

    void loadPlugins(const QString& directory = QString());
    void unloadPlugins();

    QList<PluginInfo> plugins() const;
    int pluginCount() const;
    bool hasPlugin(const QString& name) const;

    QStringList pluginNames() const;
    QStringList loadedPluginNames() const;

    PluginContext* context() const;
    void setContext(PluginContext* context);

signals:
    void pluginLoaded(const QString& name);
    void pluginUnloaded(const QString& name);
    void pluginError(const QString& name, const QString& error);

private:
    QList<PluginInfo> m_plugins;
    PluginContext* m_context = nullptr;
    QString m_directory;

    void loadPlugin(const QString& filePath);
};
