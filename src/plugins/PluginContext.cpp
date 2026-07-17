#include "PluginContext.h"

PluginContext::PluginContext(QObject *parent)
    : QObject(parent)
{
}

MainWindow *PluginContext::mainWindow() const
{
    return m_mainWindow;
}

void PluginContext::setMainWindow(MainWindow *mw)
{
    m_mainWindow = mw;
}

DocumentManager *PluginContext::documentManager() const
{
    return m_documentManager;
}

void PluginContext::setDocumentManager(DocumentManager *dm)
{
    m_documentManager = dm;
}

Settings *PluginContext::settings() const
{
    return m_settings;
}

void PluginContext::setSettings(Settings *s)
{
    m_settings = s;
}

PluginManager *PluginContext::pluginManager() const
{
    return m_pluginManager;
}

void PluginContext::setPluginManager(PluginManager *pm)
{
    m_pluginManager = pm;
    emit contextReady();
}
