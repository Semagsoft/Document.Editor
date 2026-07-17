#pragma once

#include <QString>
#include <QtPlugin>

class PluginContext;

class IPlugin
{
public:
    virtual ~IPlugin() = default;

    virtual QString name() const = 0;
    virtual QString version() const = 0;
    virtual QString description() const = 0;
    virtual bool initialize(PluginContext *context) = 0;
    virtual void shutdown() = 0;
};

#define IPlugin_IID "org.semagsoft.DocumentEditor.IPlugin"
Q_DECLARE_INTERFACE(IPlugin, IPlugin_IID)
