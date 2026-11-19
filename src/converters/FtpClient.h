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

    // Download completion handshake: the transfer is only finalized once the
    // data socket has closed AND the control channel reported 226/2xx.
    bool m_dataTransferDone = false;
    bool m_controlTransferDone = false;
    bool m_finalized = false;

    // Buffer for reading
    QByteArray m_replyBuffer;
    qint64 m_expectedBytes = 0;
    QByteArray m_dataBuffer;
    qint64 m_lastProgressBytes = 0;

    // Timeout: the global timer is an absolute deadline for the whole
    // transfer and is never renewed by incoming data. The idle timer is
    // restarted on each received chunk and aborts stalled transfers.
    QTimer *m_timeoutTimer = nullptr;
    QTimer *m_idleTimer = nullptr;
    QElapsedTimer m_elapsedTimer;
    static constexpr int kGlobalTimeoutMs = 600000; // 10 minutes, absolute
    static constexpr int kIdleTimeoutMs = 60000;    // no data for 60s aborts
    static constexpr qint64 kMaxDownloadBytes = 256LL * 1024 * 1024;

    bool parseUrl(const QString &url, QString &host, quint16 &port,
                  QString &path, QString &user, QString &password,
                  QString &error) const;
    void sendCommand(const QString &cmd);
    void connectDataChannel();
    void startTransfer();
    void finishTransfer();
    void tryFinalizeDownload();
    bool writeDownloadedData(QString &error);
    void abortTransfer(const QString &error);

    void onControlReadyRead();
    void onDataReadyRead();
    void onDataFinished();
    void onSocketError(QAbstractSocket::SocketError error);
    void onDataError(QAbstractSocket::SocketError error);
};
