#pragma once

#include <QObject>
#include <QtPlugin>
#include "plugins/IPlugin.h"

class ExamplePlugin : public QObject, public IPlugin
{
    Q_OBJECT
    Q_PLUGIN_METADATA(IID "org.semagsoft.DocumentEditor.IPlugin" FILE "exampleplugin.json")
    Q_INTERFACES(IPlugin)

public:
    QString name() const override;
    QString version() const override;
    QString description() const override;
    bool initialize(PluginContext *context) override;
    void shutdown() override;
};
