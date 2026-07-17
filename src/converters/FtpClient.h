#pragma once

#include <QObject>
#include <QString>
#include <QTcpSocket>
#include <QTimer>
#include <QElapsedTimer>

class FtpClient : public QObject
{
    Q_OBJECT

public:
    explicit FtpClient(QObject *parent = nullptr);
    ~FtpClient() override;

    void upload(const QString &localPath, const QString &remoteUrl,
                const QString &username = QString(),
                const QString &password = QString());
    void download(const QString &remoteUrl, const QString &localPath,
                  const QString &username = QString(),
                  const QString &password = QString());

    bool isBusy() const;
    void cancel();

signals:
    void uploadCompleted();
    void downloadCompleted();
    void transferFailed(const QString &error);
    void transferProgress(qint64 bytesSent, qint64 bytesTotal);

private:
    enum class State { Idle, Connecting, Connected, LoggedIn, Passive, Transferring };

    QTcpSocket *m_control = nullptr;
    QTcpSocket *m_data = nullptr;
    State m_state = State::Idle;

    QString m_host;
    quint16 m_port = 21;
    QString m_remotePath;
    QString m_localPath;
    QString m_username;
    QString m_password;
    bool m_isUpload = false;

    // PASV data
    QString m_dataHost;
    quint16 m_dataPort = 0;

    // Buffer for reading
    QByteArray m_replyBuffer;
    qint64 m_expectedBytes = 0;
    QByteArray m_dataBuffer;
    qint64 m_lastProgressBytes = 0;

    // Timeout
    QTimer *m_timeoutTimer = nullptr;
    QElapsedTimer m_elapsedTimer;
    static constexpr int kGlobalTimeoutMs = 60000;

    void parseUrl(const QString &url, QString &host, quint16 &port,
                  QString &path) const;
    void sendCommand(const QString &cmd);
    void connectDataChannel();
    void startTransfer();
    void abortTransfer(const QString &error);

    void onControlReadyRead();
    void onDataReadyRead();
    void onDataFinished();
    void onSocketError(QAbstractSocket::SocketError error);
};
