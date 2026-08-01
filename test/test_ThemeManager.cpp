#include <QTest>
#include <QSignalSpy>
#include <QApplication>
#include "mainwindow/ThemeManager.h"

class TestThemeManager : public QObject
{
    Q_OBJECT

private slots:
    void testDefaultTheme()
    {
        ThemeManager tm;
        QCOMPARE(tm.currentTheme(), ThemeManager::Office2010Blue);
    }

    void testSetTheme()
    {
        ThemeManager tm;
        tm.setTheme(ThemeManager::Windows8);
        QCOMPARE(tm.currentTheme(), ThemeManager::Windows8);

        tm.setTheme(ThemeManager::Office2010Black);
        QCOMPARE(tm.currentTheme(), ThemeManager::Office2010Black);
    }

    void testThemeChangedSignal()
    {
        ThemeManager tm;
        QSignalSpy spy(&tm, &ThemeManager::themeChanged);

        tm.setTheme(ThemeManager::Office2013);
        QCOMPARE(spy.count(), 1);
        QCOMPARE(spy.takeFirst().at(0).value<ThemeManager::Theme>(),
                 ThemeManager::Office2013);

        tm.setTheme(ThemeManager::Office2010Blue);
        QCOMPARE(spy.count(), 1);
    }

    void testStylesheetResetOnMissingResource()
    {
        ThemeManager tm;
        qApp->setStyleSheet(QStringLiteral("QMainWindow { background-color: #ff00ff; }"));
        QVERIFY(!qApp->styleSheet().isEmpty());

        tm.setTheme(ThemeManager::Office2010Blue);
        QVERIFY(qApp->styleSheet().isEmpty());
    }
};

QTEST_MAIN(TestThemeManager)
#include "test_ThemeManager.moc"
