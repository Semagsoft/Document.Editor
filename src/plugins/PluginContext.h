#pragma once

#include <QObject>

class MainWindow;
class DocumentManager;
class Settings;
class PluginManager;

class PluginContext : public QObject
{
    Q_OBJECT

public:
    explicit PluginContext(QObject *parent = nullptr);

    MainWindow *mainWindow() const;
    void setMainWindow(MainWindow *mw);

    DocumentManager *documentManager() const;
    void setDocumentManager(DocumentManager *dm);

    Settings *settings() const;
    void setSettings(Settings *s);

    PluginManager *pluginManager() const;
    void setPluginManager(PluginManager *pm);

signals:
    void contextReady();

private:
    MainWindow *m_mainWindow = nullptr;
    DocumentManager *m_documentManager = nullptr;
    Settings *m_settings = nullptr;
    PluginManager *m_pluginManager = nullptr;
};
