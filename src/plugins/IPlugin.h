#pragma once

#include <QString>
#include <QtPlugin>

class PluginContext;

class IPlugin
{
public:
    // Bumped whenever the plugin API/ABI changes. New plugins must declare
    // "apiVersion": <this value> in their metadata JSON; PluginManager
    // refuses plugins whose apiVersion does not match, before any plugin
    // code runs. Plugins compiled against a newer header report it via
    // interfaceVersion() as a second line of defense.
    static constexpr int InterfaceVersion = 1;

    virtual ~IPlugin() = default;

    virtual QString name() const = 0;
    virtual QString version() const = 0;
    virtual QString description() const = 0;
    virtual bool initialize(PluginContext *context) = 0;
    virtual void shutdown() = 0;
    virtual int interfaceVersion() const { return InterfaceVersion; }
};

#define IPlugin_IID "org.semagsoft.DocumentEditor.IPlugin"
Q_DECLARE_INTERFACE(IPlugin, IPlugin_IID)
