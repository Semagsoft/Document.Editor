#pragma once

#include <QObject>
#include <QString>

class ThemeManager : public QObject
{
    Q_OBJECT

public:
    enum Theme {
        Office2010Blue,
        Office2010Silver,
        Office2010Black,
        Office2013,
        Windows8
    };

    explicit ThemeManager(QObject *parent = nullptr);

    Theme currentTheme() const;
    void setTheme(Theme theme);

signals:
    void themeChanged(Theme theme);

private:
    void setPalette(Theme theme) const;
    QString loadStylesheet(Theme theme) const;
    Theme m_currentTheme = Office2010Blue;
};
