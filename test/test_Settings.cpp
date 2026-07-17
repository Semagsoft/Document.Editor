#include <QTest>
#include <QSignalSpy>
#include <QFont>
#include <QRect>
#include <QSettings>
#include "app/Settings.h"

class TestSettings : public QObject
{
    Q_OBJECT

private slots:
    void initTestCase()
    {
        QSettings existing(QSettings::IniFormat, QSettings::UserScope,
            QStringLiteral("Semagsoft"), QStringLiteral("Document.Editor"));
        existing.clear();
        existing.sync();
    }

    void testDefaults()
    {
        Settings s;
        QCOMPARE(s.startupMode(), 0);
        QCOMPARE(s.showStartupDialog(), true);
        QCOMPARE(s.checkForUpdatesOnStartup(), true);
        QCOMPARE(s.theme(), 0);
        QCOMPARE(s.enableGlass(), false);
        QCOMPARE(s.defaultFontSize(), 12);
        QCOMPARE(s.spellCheckEnabled(), false);
        QCOMPARE(s.tabPlacement(), 0);
        QCOMPARE(s.tabSizeMode(), 0);
        QCOMPARE(s.showRecentDocuments(), true);
        QCOMPARE(s.rulerMeasurement(), 0);
        QCOMPARE(s.ttsVoice(), 0);
        QCOMPARE(s.ttsSpeed(), 0);
        QCOMPARE(s.pluginsEnabled(), false);
        QCOMPARE(s.windowMaximized(), false);
        QVERIFY(s.recentFiles().isEmpty());
    }

    void testSetGet()
    {
        Settings s;
        QSignalSpy spy(&s, &Settings::settingsChanged);

        s.setStartupMode(1);
        QCOMPARE(s.startupMode(), 1);
        QCOMPARE(spy.count(), 1);

        s.setTheme(3);
        QCOMPARE(s.theme(), 3);
        QCOMPARE(spy.count(), 2);

        s.setShowRuler(false);
        QCOMPARE(s.showRuler(), false);
        QCOMPARE(spy.count(), 3);

        s.setDefaultFontSize(14);
        QCOMPARE(s.defaultFontSize(), 14);
        QCOMPARE(spy.count(), 4);

        s.setPluginsEnabled(true);
        QCOMPARE(s.pluginsEnabled(), true);
        QCOMPARE(spy.count(), 5);
    }

    void testRecentFiles()
    {
        Settings s;
        QVERIFY(s.recentFiles().isEmpty());

        s.addRecentFile(QStringLiteral("doc1.xaml"));
        QCOMPARE(s.recentFiles().size(), 1);
        QCOMPARE(s.recentFiles().first(), QStringLiteral("doc1.xaml"));

        s.addRecentFile(QStringLiteral("doc2.xaml"));
        QCOMPARE(s.recentFiles().size(), 2);
        QCOMPARE(s.recentFiles().first(), QStringLiteral("doc2.xaml"));

        s.addRecentFile(QStringLiteral("doc1.xaml"));
        QCOMPARE(s.recentFiles().size(), 2);
        QCOMPARE(s.recentFiles().first(), QStringLiteral("doc1.xaml"));

        s.clearRecentFiles();
        QVERIFY(s.recentFiles().isEmpty());
    }

    void testWindowGeometry()
    {
        Settings s;
        QSignalSpy spy(&s, &Settings::settingsChanged);

        QRect geom(50, 50, 800, 600);
        s.setWindowGeometry(geom);
        QCOMPARE(s.windowGeometry(), geom);
        QCOMPARE(spy.count(), 1);

        s.setWindowMaximized(true);
        QCOMPARE(s.windowMaximized(), true);
        QCOMPARE(spy.count(), 2);
    }
};

QTEST_MAIN(TestSettings)
#include "test_Settings.moc"
