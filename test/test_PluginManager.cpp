#include <QTest>
#include <QSignalSpy>
#include <QDir>
#include <QTemporaryDir>
#include "plugins/PluginManager.h"
#include "plugins/PluginContext.h"
#include "plugins/IPlugin.h"

class TestPluginManager : public QObject
{
    Q_OBJECT

private:
    PluginManager *m_manager = nullptr;
    PluginContext *m_context = nullptr;

private slots:
    void initTestCase()
    {
        m_manager = new PluginManager(this);
        m_context = new PluginContext(this);
        m_manager->setContext(m_context);
        m_context->setPluginManager(m_manager);
    }

    void testInitialState()
    {
        QCOMPARE(m_manager->pluginCount(), 0);
        QVERIFY(m_manager->plugins().isEmpty());
        QVERIFY(m_manager->pluginNames().isEmpty());
        QVERIFY(m_manager->loadedPluginNames().isEmpty());
    }

    void testNoPluginsInitially()
    {
        QVERIFY(!m_manager->hasPlugin(QStringLiteral("nonexistent")));
        QCOMPARE(m_manager->pluginNames().size(), 0);
    }

    void testLoadFromEmptyDirectory()
    {
        // Loading from a directory with no shared libraries should be a no-op
        QTemporaryDir tempDir;
        QVERIFY(tempDir.isValid());

        m_manager->loadPlugins(tempDir.path());
        QCOMPARE(m_manager->pluginCount(), 0);
    }

    void testLoadFromNonexistentDirectory()
    {
        // Loading from a directory that doesn't exist should be a no-op
        m_manager->loadPlugins(QStringLiteral("/nonexistent/plugins/directory"));
        QCOMPARE(m_manager->pluginCount(), 0);
    }

    void testUnloadWithNoPlugins()
    {
        // Unloading with no plugins should not crash
        m_manager->unloadPlugins();
        QCOMPARE(m_manager->pluginCount(), 0);
    }

    void testPluginSignals()
    {
        QSignalSpy loadSpy(m_manager, &PluginManager::pluginLoaded);
        QSignalSpy unloadSpy(m_manager, &PluginManager::pluginUnloaded);
        QSignalSpy errorSpy(m_manager, &PluginManager::pluginError);

        Q_UNUSED(loadSpy);
        Q_UNUSED(unloadSpy);
        Q_UNUSED(errorSpy);

        // Verify signals are accessible
        QVERIFY(connect(m_manager, &PluginManager::pluginLoaded, this, [](const QString &) {}));
        QVERIFY(connect(m_manager, &PluginManager::pluginUnloaded, this, [](const QString &) {}));
        QVERIFY(connect(m_manager, &PluginManager::pluginError, this, [](const QString &, const QString &) {}));
    }

    void testMultipleLoadCalls()
    {
        // Calling load multiple times should not crash
        QTemporaryDir tempDir;
        QVERIFY(tempDir.isValid());

        m_manager->loadPlugins(tempDir.path());
        m_manager->loadPlugins(tempDir.path());
        m_manager->loadPlugins(tempDir.path());
        // No crash = success
        QVERIFY(true);
    }

    void testContextIsAccessible()
    {
        PluginContext *ctx = m_manager->context();
        QVERIFY(ctx != nullptr);
        QCOMPARE(ctx, m_context);
    }
};

QTEST_MAIN(TestPluginManager)
#include "test_PluginManager.moc"
