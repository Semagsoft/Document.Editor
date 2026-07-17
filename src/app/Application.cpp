#include "Application.h"
#include "Settings.h"

#include <QStyleFactory>

Application::Application(int &argc, char **argv)
    : QApplication(argc, argv)
{
    setApplicationName(QStringLiteral("Document.Editor"));
    setOrganizationName(QStringLiteral("Semagsoft"));
    setApplicationVersion(QStringLiteral(APP_VERSION));
    setOrganizationDomain(QStringLiteral("documenteditor.net"));

    setStyle(QStyleFactory::create(QStringLiteral("Fusion")));

    m_settings = new Settings(this);

    for (int i = 1; i < argc; ++i)
        m_startupFiles.append(QString::fromLocal8Bit(argv[i]));
}

Application::~Application()
{
    if (m_settings)
        m_settings->save();
}

Settings *Application::settings() const
{
    return m_settings;
}

QStringList Application::startupFiles() const
{
    return m_startupFiles;
}
