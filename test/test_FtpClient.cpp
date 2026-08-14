#include <QTest>
#include <QTcpServer>
#include <QTcpSocket>
#include <QHostAddress>
#include <QSignalSpy>
#include <QTemporaryDir>
#include "converters/FtpClient.h"

// Minimal passive-mode FTP server sufficient to drive FtpClient through real
// upload and download transfers. Runs inside the test's event loop.
class FakeFtpServer : public QObject
{
    Q_OBJECT

public:
    QTcpServer controlServer;
    QTcpServer dataServer;

    QTcpSocket *control = nullptr;
    QTcpSocket *data = nullptr;

    QByteArray retrPayload = QByteArrayLiteral("hello ftp payload");
    QByteArray uploaded;
    QString mode;
    bool retrPending = false;
    int completedTransfers = 0;

    bool listen()
    {
        return controlServer.listen(QHostAddress::LocalHost, 0);
    }

    quint16 port() const { return controlServer.serverPort(); }

    void start()
    {
        connect(&controlServer, &QTcpServer::newConnection, this, [this]() {
            QTcpSocket *s = controlServer.nextPendingConnection();
            s->setParent(this);
            connect(s, &QTcpSocket::readyRead, this, &FakeFtpServer::onControlReady);
            connect(s, &QTcpSocket::disconnected, this, [this, s]() {
                if (control == s)
                    control = nullptr;
            });
            control = s;
            sendReply(QStringLiteral("220 welcome"));
        });

        connect(&dataServer, &QTcpServer::newConnection, this, [this]() {
            QTcpSocket *s = dataServer.nextPendingConnection();
            s->setParent(this);
            data = s;
            // Accumulate any uploaded payload and complete the transfer when the
            // peer closes the data channel. The STOR/RETR command arrives on the
            // control connection *after* this data connection is established, so
            // the wiring must not depend on mode.
            connect(s, &QTcpSocket::readyRead, this, [this, s]() {
                if (data == s)
                    uploaded += s->readAll();
            });
            connect(s, &QTcpSocket::disconnected, this, [this, s]() {
                if (data != s)
                    return;
                sendReply(QStringLiteral("226 transfer complete"));
                ++completedTransfers;
            });
            if (mode == QStringLiteral("retr"))
                tryRetr();
        });
    }

    void sendReply(const QString &reply)
    {
        if (control)
            control->write((reply + QStringLiteral("\r\n")).toUtf8());
    }

private:
    void tryRetr()
    {
        if (!retrPending)
            return;
        if (!data)
            return;
        retrPending = false;
        data->write(retrPayload);
        data->disconnectFromHost();
    }

    void onControlReady()
    {
        while (control && control->canReadLine()) {
            const QByteArray line = control->readLine().trimmed().toUpper();
            handle(line);
        }
    }

    void handle(const QByteArray &line)
    {
        if (line == QByteArrayLiteral("USER ANONYMOUS")) {
            sendReply(QStringLiteral("331 need password"));
        } else if (line.startsWith(QByteArrayLiteral("PASS"))) {
            sendReply(QStringLiteral("230 logged in"));
        } else if (line == QByteArrayLiteral("PASV")) {
            dataServer.close();
            dataServer.listen(QHostAddress::LocalHost, 0);
            const quint16 p = dataServer.serverPort();
            sendReply(QStringLiteral("227 Entering Passive Mode (127,0,0,1,%1,%2)")
                          .arg(p / 256).arg(p % 256));
        } else if (line.startsWith(QByteArrayLiteral("TYPE"))) {
            sendReply(QStringLiteral("200 type set"));
        } else if (line.startsWith(QByteArrayLiteral("RETR"))) {
            mode = QStringLiteral("retr");
            retrPending = true;
            sendReply(QStringLiteral("150 opening data"));
            tryRetr();
        } else if (line.startsWith(QByteArrayLiteral("STOR"))) {
            mode = QStringLiteral("stor");
            sendReply(QStringLiteral("150 opening data"));
        } else if (line == QByteArrayLiteral("QUIT")) {
            sendReply(QStringLiteral("221 goodbye"));
        } else {
            sendReply(QStringLiteral("200 ok"));
        }
    }
};

class TestFtpClient : public QObject
{
    Q_OBJECT
    FakeFtpServer m_server;

private slots:
    void initTestCase()
    {
        QVERIFY(m_server.listen());
        m_server.start();
    }

    void cleanupTestCase()
    {
        m_server.controlServer.close();
        m_server.dataServer.close();
    }

    QString remoteUrl(const QString &file) const
    {
        return QStringLiteral("ftp://127.0.0.1:%1/%2")
            .arg(m_server.port()).arg(file);
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
        client.upload(QStringLiteral("/tmp/does_not_exist.txt"), remoteUrl(QStringLiteral("x.txt")));
        QVERIFY(client.isBusy());
        client.upload(QStringLiteral("/tmp/test2.txt"), remoteUrl(QStringLiteral("x2.txt")));
        client.cancel();
        QVERIFY(!client.isBusy());
    }

    void testSignalsExist()
    {
        FtpClient client;
        QVERIFY(connect(&client, &FtpClient::uploadCompleted, this, []() {}));
        QVERIFY(connect(&client, &FtpClient::downloadCompleted, this, []() {}));
        QVERIFY(connect(&client, &FtpClient::transferFailed, this, [](const QString &) {}));
        QVERIFY(connect(&client, &FtpClient::transferProgress, this, [](qint64, qint64) {}));
    }

    void testDownloadAndReuse()
    {
        QTemporaryDir dir;
        QVERIFY(dir.isValid());
        m_server.completedTransfers = 0;
        FtpClient client;
        QSignalSpy doneSpy(&client, &FtpClient::downloadCompleted);

        // First transfer.
        client.download(remoteUrl(QStringLiteral("a.bin")),
                        dir.filePath(QStringLiteral("out1.bin")));
        QTRY_COMPARE_WITH_TIMEOUT(doneSpy.count(), 1, 5000);
        QCOMPARE(doneSpy.count(), 1);
        // The data socket must have fully written the payload before the
        // completion signal was emitted.
        QFile f1(dir.filePath(QStringLiteral("out1.bin")));
        QVERIFY(f1.open(QIODevice::ReadOnly));
        QCOMPARE(f1.readAll(), m_server.retrPayload);
        f1.close();

        // Second transfer on the same instance must not hang (regression test
        // for the old m_data->disconnect() that tore down the data wiring).
        client.download(remoteUrl(QStringLiteral("b.bin")),
                        dir.filePath(QStringLiteral("out2.bin")));
        QTRY_COMPARE_WITH_TIMEOUT(doneSpy.count(), 2, 5000);
        QFile f2(dir.filePath(QStringLiteral("out2.bin")));
        QVERIFY(f2.open(QIODevice::ReadOnly));
        QCOMPARE(f2.readAll(), m_server.retrPayload);
        f2.close();

        QCOMPARE(m_server.completedTransfers, 2);
        QVERIFY(!client.isBusy());
    }

    void testUploadAndReuse()
    {
        QTemporaryDir dir;
        QVERIFY(dir.isValid());
        m_server.completedTransfers = 0;
        m_server.uploaded.clear();
        const QByteArray payload = QByteArrayLiteral("upload payload 12345");
        const QString src = dir.filePath(QStringLiteral("src.bin"));
        QFile w(src);
        QVERIFY(w.open(QIODevice::WriteOnly));
        w.write(payload);
        w.close();

        FtpClient client;
        QSignalSpy doneSpy(&client, &FtpClient::uploadCompleted);

        client.upload(src, remoteUrl(QStringLiteral("up1.bin")));
        QTRY_COMPARE_WITH_TIMEOUT(doneSpy.count(), 1, 5000);
        QCOMPARE(m_server.uploaded, payload);

        m_server.uploaded.clear();
        client.upload(src, remoteUrl(QStringLiteral("up2.bin")));
        QTRY_COMPARE_WITH_TIMEOUT(doneSpy.count(), 2, 5000);
        QCOMPARE(m_server.uploaded, payload);

        QCOMPARE(m_server.completedTransfers, 2);
        QVERIFY(!client.isBusy());
    }
};

QTEST_MAIN(TestFtpClient)
#include "test_FtpClient.moc"
