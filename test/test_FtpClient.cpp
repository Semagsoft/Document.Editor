#include <QTest>
#include <QTcpServer>
#include <QHostAddress>
#include "converters/FtpClient.h"

class TestFtpClient : public QObject
{
    Q_OBJECT
    QTcpServer m_server;
    quint16 m_port = 0;

private slots:
    void initTestCase()
    {
        // Bind a local listener so the client connects to localhost instead of
        // resolving external hosts (keeps tests hermetic and fast).
        QVERIFY(m_server.listen(QHostAddress::LocalHost, 0));
        m_port = m_server.serverPort();
    }

    void cleanupTestCase()
    {
        m_server.close();
    }

    QString remoteUrl(const QString &file) const
    {
        return QStringLiteral("ftp://127.0.0.1:%1/%2").arg(m_port).arg(file);
    }

    void testInitialState()
    {
        FtpClient client;
        QVERIFY(!client.isBusy());
    }

    void testCancelOnIdle()
    {
        FtpClient client;
        client.cancel();
        QVERIFY(!client.isBusy());
    }

    void testDoubleUploadGetsIgnored()
    {
        FtpClient client;
        client.upload(QStringLiteral("/tmp/test.txt"), remoteUrl(QStringLiteral("test.txt")));
        QVERIFY(client.isBusy());

        // Second upload while busy should be safely ignored
        client.upload(QStringLiteral("/tmp/test2.txt"), remoteUrl(QStringLiteral("test2.txt")));

        client.cancel();
        QVERIFY(!client.isBusy());
    }

    void testDoubleDownloadGetsIgnored()
    {
        FtpClient client;
        client.download(remoteUrl(QStringLiteral("test.txt")), QStringLiteral("/tmp/test.txt"));
        QVERIFY(client.isBusy());

        client.download(remoteUrl(QStringLiteral("test2.txt")), QStringLiteral("/tmp/test2.txt"));

        client.cancel();
        QVERIFY(!client.isBusy());
    }

    void testFtpClientReuse()
    {
        FtpClient client;

        client.upload(QStringLiteral("/tmp/test.txt"), remoteUrl(QStringLiteral("test.txt")));
        client.cancel();
        QVERIFY(!client.isBusy());

        // Can reuse after cancel
        client.download(remoteUrl(QStringLiteral("test.txt")), QStringLiteral("/tmp/test.txt"));
        client.cancel();
        QVERIFY(!client.isBusy());
    }

    void testSignalsExist()
    {
        FtpClient client;
        // Verify signals are accessible (compile-time check)
        QVERIFY(connect(&client, &FtpClient::uploadCompleted, this, []() {}));
        QVERIFY(connect(&client, &FtpClient::downloadCompleted, this, []() {}));
        QVERIFY(connect(&client, &FtpClient::transferFailed, this, [](const QString &) {}));
        QVERIFY(connect(&client, &FtpClient::transferProgress, this, [](qint64, qint64) {}));
    }
};

QTEST_MAIN(TestFtpClient)
#include "test_FtpClient.moc"
