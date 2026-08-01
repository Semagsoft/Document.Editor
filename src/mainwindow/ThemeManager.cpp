#include "ThemeManager.h"

#include <QApplication>
#include <QPalette>
#include <QStyleFactory>
#include <QStyle>
#include <QFile>
#include <QTextStream>

ThemeManager::ThemeManager(QObject *parent)
    : QObject(parent)
{
    setTheme(Office2010Blue);
}

ThemeManager::Theme ThemeManager::currentTheme() const
{
    return m_currentTheme;
}

void ThemeManager::setTheme(Theme theme)
{
    m_currentTheme = theme;

    setPalette(theme);

    QString stylesheet = loadStylesheet(theme);
    qApp->setStyleSheet(stylesheet);

    emit themeChanged(theme);
}

void ThemeManager::setPalette(Theme theme) const
{
    QPalette pal;
    switch (theme) {
    case Office2010Blue:
        pal.setColor(QPalette::Window, QColor(222, 234, 246));
        pal.setColor(QPalette::WindowText, Qt::black);
        pal.setColor(QPalette::Base, Qt::white);
        pal.setColor(QPalette::AlternateBase, QColor(222, 234, 246));
        pal.setColor(QPalette::ToolTipBase, QColor(255, 255, 220));
        pal.setColor(QPalette::ToolTipText, Qt::black);
        pal.setColor(QPalette::Text, Qt::black);
        pal.setColor(QPalette::Button, QColor(222, 234, 246));
        pal.setColor(QPalette::ButtonText, Qt::black);
        pal.setColor(QPalette::BrightText, Qt::red);
        pal.setColor(QPalette::Link, QColor(0, 102, 204));
        pal.setColor(QPalette::Highlight, QColor(0, 102, 204));
        pal.setColor(QPalette::HighlightedText, Qt::white);
        break;
    case Office2010Black:
        pal.setColor(QPalette::Window, QColor(80, 80, 80));
        pal.setColor(QPalette::WindowText, Qt::white);
        pal.setColor(QPalette::Base, QColor(60, 60, 60));
        pal.setColor(QPalette::AlternateBase, QColor(80, 80, 80));
        pal.setColor(QPalette::ToolTipBase, QColor(255, 255, 220));
        pal.setColor(QPalette::ToolTipText, Qt::black);
        pal.setColor(QPalette::Text, Qt::white);
        pal.setColor(QPalette::Button, QColor(80, 80, 80));
        pal.setColor(QPalette::ButtonText, Qt::white);
        pal.setColor(QPalette::BrightText, Qt::red);
        pal.setColor(QPalette::Link, QColor(102, 153, 255));
        pal.setColor(QPalette::Highlight, QColor(102, 153, 255));
        pal.setColor(QPalette::HighlightedText, Qt::black);
        break;
    case Office2010Silver:
        pal.setColor(QPalette::Window, QColor(196, 199, 207));
        pal.setColor(QPalette::WindowText, Qt::black);
        pal.setColor(QPalette::Base, Qt::white);
        pal.setColor(QPalette::AlternateBase, QColor(196, 199, 207));
        pal.setColor(QPalette::ToolTipBase, QColor(255, 255, 220));
        pal.setColor(QPalette::ToolTipText, Qt::black);
        pal.setColor(QPalette::Text, Qt::black);
        pal.setColor(QPalette::Button, QColor(196, 199, 207));
        pal.setColor(QPalette::ButtonText, Qt::black);
        pal.setColor(QPalette::BrightText, Qt::red);
        pal.setColor(QPalette::Link, QColor(0, 102, 204));
        pal.setColor(QPalette::Highlight, QColor(84, 106, 141));
        pal.setColor(QPalette::HighlightedText, Qt::white);
        break;
    case Office2013:
        pal.setColor(QPalette::Window, QColor(255, 255, 255));
        pal.setColor(QPalette::WindowText, QColor(68, 68, 68));
        pal.setColor(QPalette::Base, Qt::white);
        pal.setColor(QPalette::AlternateBase, QColor(241, 241, 241));
        pal.setColor(QPalette::ToolTipBase, QColor(255, 255, 220));
        pal.setColor(QPalette::ToolTipText, Qt::black);
        pal.setColor(QPalette::Text, QColor(68, 68, 68));
        pal.setColor(QPalette::Button, QColor(225, 225, 225));
        pal.setColor(QPalette::ButtonText, QColor(68, 68, 68));
        pal.setColor(QPalette::BrightText, Qt::red);
        pal.setColor(QPalette::Link, QColor(0, 102, 204));
        pal.setColor(QPalette::Highlight, QColor(0, 114, 198));
        pal.setColor(QPalette::HighlightedText, Qt::white);
        break;
    case Windows8:
        pal.setColor(QPalette::Window, QColor(241, 241, 241));
        pal.setColor(QPalette::WindowText, QColor(30, 30, 30));
        pal.setColor(QPalette::Base, Qt::white);
        pal.setColor(QPalette::AlternateBase, QColor(241, 241, 241));
        pal.setColor(QPalette::ToolTipBase, QColor(255, 255, 220));
        pal.setColor(QPalette::ToolTipText, Qt::black);
        pal.setColor(QPalette::Text, QColor(30, 30, 30));
        pal.setColor(QPalette::Button, QColor(225, 225, 225));
        pal.setColor(QPalette::ButtonText, QColor(30, 30, 30));
        pal.setColor(QPalette::BrightText, Qt::red);
        pal.setColor(QPalette::Link, QColor(0, 102, 204));
        pal.setColor(QPalette::Highlight, QColor(0, 114, 198));
        pal.setColor(QPalette::HighlightedText, Qt::white);
        break;
    }
    qApp->setPalette(pal);
}

QString ThemeManager::loadStylesheet(Theme theme) const
{
    QString path;
    switch (theme) {
    case Office2010Blue:   path = QStringLiteral(":/styles/office2010blue.qss"); break;
    case Office2010Silver: path = QStringLiteral(":/styles/office2010silver.qss"); break;
    case Office2010Black:  path = QStringLiteral(":/styles/office2010black.qss"); break;
    case Office2013:       path = QStringLiteral(":/styles/office2013.qss"); break;
    case Windows8:         path = QStringLiteral(":/styles/windows8.qss"); break;
    }

    QFile file(path);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
        return {};

    QTextStream stream(&file);
    return stream.readAll();
}
