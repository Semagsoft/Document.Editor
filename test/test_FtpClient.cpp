#include <QTest>
#include "converters/FtpClient.h"

class TestFtpClient : public QObject
{
    Q_OBJECT

private slots:
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
        client.upload(QStringLiteral("/tmp/test.txt"),
                       QStringLiteral("ftp://example.com/test.txt"));
        QVERIFY(client.isBusy());

        // Second upload while busy should be safely ignored
        client.upload(QStringLiteral("/tmp/test2.txt"),
                       QStringLiteral("ftp://example.com/test2.txt"));

        client.cancel();
        QVERIFY(!client.isBusy());
    }

    void testDoubleDownloadGetsIgnored()
    {
        FtpClient client;
        client.download(QStringLiteral("ftp://example.com/test.txt"),
                         QStringLiteral("/tmp/test.txt"));
        QVERIFY(client.isBusy());

        client.download(QStringLiteral("ftp://example.com/test2.txt"),
                         QStringLiteral("/tmp/test2.txt"));

        client.cancel();
        QVERIFY(!client.isBusy());
    }

    void testFtpClientReuse()
    {
        FtpClient client;

        client.upload(QStringLiteral("/tmp/test.txt"),
                       QStringLiteral("ftp://example.com/test.txt"));
        client.cancel();
        QVERIFY(!client.isBusy());

        // Can reuse after cancel
        client.download(QStringLiteral("ftp://example.com/test.txt"),
                         QStringLiteral("/tmp/test.txt"));
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
