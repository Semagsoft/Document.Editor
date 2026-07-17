#include "ExamplePlugin.h"
#include "plugins/PluginContext.h"

#include <QDebug>

QString ExamplePlugin::name() const
{
    return QStringLiteral("ExamplePlugin");
}

QString ExamplePlugin::version() const
{
    return QStringLiteral("1.0.0");
}

QString ExamplePlugin::description() const
{
    return QStringLiteral("An example Document.Editor plugin");
}

bool ExamplePlugin::initialize(PluginContext *context)
{
    Q_UNUSED(context)
    qDebug() << "ExamplePlugin: initialized successfully";
    return true;
}

void ExamplePlugin::shutdown()
{
    qDebug() << "ExamplePlugin: shutting down";
}
