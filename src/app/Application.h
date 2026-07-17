#pragma once

#include <QApplication>
#include <QStringList>

class Settings;

class Application : public QApplication
{
    Q_OBJECT

public:
    Application(int &argc, char **argv);
    ~Application() override;

    Settings *settings() const;

    QStringList startupFiles() const;

private:
    Settings *m_settings = nullptr;
    QStringList m_startupFiles;
};
