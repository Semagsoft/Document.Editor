#include "FtpClient.h"

#include <QFile>
#include <QFileInfo>
#include <QRegularExpression>
#include <QUrl>

FtpClient::FtpClient(QObject* parent)
    : QObject(parent)
{
    m_control = new QTcpSocket(this);
    connect(m_control, &QTcpSocket::readyRead, this, &FtpClient::onControlReadyRead);
    connect(m_control, &QTcpSocket::errorOccurred, this, &FtpClient::onSocketError);

    m_data = new QTcpSocket(this);
    connect(m_data, &QTcpSocket::readyRead, this, &FtpClient::onDataReadyRead);
    connect(m_data, &QTcpSocket::disconnected, this, &FtpClient::onDataFinished);

    m_timeoutTimer = new QTimer(this);
    m_timeoutTimer->setSingleShot(true);
    connect(m_timeoutTimer, &QTimer::timeout, this, [this]() {
        abortTransfer(tr("FTP transfer timed out after %1 seconds")
                .arg(kGlobalTimeoutMs / 1000));
    });
}

FtpClient::~FtpClient()
{
    cancel();
}

void FtpClient::parseUrl(const QString& url, QString& host, quint16& port,
    QString& path) const
{
    QUrl qurl(url);
    host = qurl.host();
    port = qurl.port(21);
    path = qurl.path();
    if (path.isEmpty())
        path = QStringLiteral("/");
}

void FtpClient::upload(const QString& localPath, const QString& remoteUrl,
    const QString& username, const QString& password)
{
    if (isBusy())
        return;

    m_localPath = localPath;
    parseUrl(remoteUrl, m_host, m_port, m_remotePath);
    m_username = username.isEmpty() ? QStringLiteral("anonymous") : username;
    m_password = password.isEmpty() ? QStringLiteral("anonymous@") : password;
    m_isUpload = true;

    m_state = State::Connecting;
    m_control->connectToHost(m_host, m_port);
    m_elapsedTimer.start();
    m_lastProgressBytes = 0;
    m_timeoutTimer->start(kGlobalTimeoutMs);
}

void FtpClient::download(const QString& remoteUrl, const QString& localPath,
    const QString& username, const QString& password)
{
    if (isBusy())
        return;

    m_localPath = localPath;
    parseUrl(remoteUrl, m_host, m_port, m_remotePath);
    m_username = username.isEmpty() ? QStringLiteral("anonymous") : username;
    m_password = password.isEmpty() ? QStringLiteral("anonymous@") : password;
    m_isUpload = false;

    m_state = State::Connecting;
    m_control->connectToHost(m_host, m_port);
    m_elapsedTimer.start();
    m_lastProgressBytes = 0;
    m_timeoutTimer->start(kGlobalTimeoutMs);
}

bool FtpClient::isBusy() const
{
    return m_state != State::Idle;
}

void FtpClient::cancel()
{
    if (m_data->state() != QTcpSocket::UnconnectedState)
        m_data->disconnectFromHost();
    if (m_control->state() != QTcpSocket::UnconnectedState) {
        sendCommand(QStringLiteral("QUIT"));
        m_control->disconnectFromHost();
    }
    m_state = State::Idle;
    m_replyBuffer.clear();
    m_dataBuffer.clear();
    m_timeoutTimer->stop();
}

void FtpClient::sendCommand(const QString& cmd)
{
    m_control->write((cmd + QStringLiteral("\r\n")).toUtf8());
    m_control->flush();
}

void FtpClient::onControlReadyRead()
{
    m_replyBuffer += m_control->readAll();

    // FTP replies end with "NNN <text>\r\n" for last line
    // or "NNN-<text>\r\n" for multi-line
    while (true) {
        int endIdx = m_replyBuffer.indexOf(QByteArrayLiteral("\r\n"));
        if (endIdx < 0)
            break;

        QByteArray line = m_replyBuffer.left(endIdx);
        m_replyBuffer.remove(0, endIdx + 2);

        if (line.length() < 3)
            continue;

        int code = line.left(3).toInt();
        bool isLast = (line.length() < 4 || line[3] != '-');

        switch (m_state) {
        case State::Connecting:
        case State::Idle:
            if (code == 220) {
                m_state = State::Connected;
                sendCommand(QStringLiteral("USER %1").arg(m_username));
            } else {
                abortTransfer(tr("Unexpected FTP response: %1").arg(QString::fromUtf8(line)));
            }
            break;

        case State::Connected:
            if (code == 230 || code == 331) {
                if (code == 331)
                    sendCommand(QStringLiteral("PASS %1").arg(m_password));
                else
                    m_state = State::LoggedIn;
            }
            if (code == 230)
                m_state = State::LoggedIn;
            if (m_state == State::LoggedIn)
                sendCommand(QStringLiteral("PASV"));
            break;

        case State::LoggedIn:
            if (code == 227) {
                // Parse PASV response: 227 Entering Passive Mode (h1,h2,h3,h4,p1,p2)
                QRegularExpression re(QStringLiteral("\\((\\d+),(\\d+),(\\d+),(\\d+),(\\d+),(\\d+)\\)"));
                auto match = re.match(QString::fromUtf8(line));
                if (match.hasMatch()) {
                    m_dataHost = QStringLiteral("%1.%2.%3.%4")
                                     .arg(match.captured(1), match.captured(2),
                                         match.captured(3), match.captured(4));
                    m_dataPort = match.captured(5).toInt() * 256 + match.captured(6).toInt();
                    m_state = State::Passive;
                    connectDataChannel();
                } else {
                    abortTransfer(tr("Failed to parse PASV response"));
                }
            }
            break;

        case State::Passive:
            if (code >= 200 && code < 300) {
                m_state = State::Transferring;
                startTransfer();
            }
            break;

        case State::Transferring:
            if (code >= 200 && code < 300) {
                // Transfer complete
                m_state = State::Idle;
                if (m_isUpload)
                    emit uploadCompleted();
                else
                    emit downloadCompleted();
            }
            break;
        }

        if (isLast && !isBusy())
            break;
    }
}

void FtpClient::connectDataChannel()
{
    m_data->connectToHost(m_dataHost, m_dataPort);
    connect(m_data, &QTcpSocket::connected, this, [this]() {
        if (m_isUpload)
            sendCommand(QStringLiteral("STOR %1").arg(m_remotePath));
        else
            sendCommand(QStringLiteral("RETR %1").arg(m_remotePath)); }, Qt::SingleShotConnection);
}

void FtpClient::startTransfer()
{
    if (m_isUpload) {
        QFile file(m_localPath);
        if (!file.open(QIODevice::ReadOnly)) {
            abortTransfer(tr("Cannot open local file: %1").arg(m_localPath));
            return;
        }
        QByteArray data = file.readAll();
        file.close();
        m_data->write(data);
        m_data->disconnectFromHost();
        m_lastProgressBytes = data.size();
        m_timeoutTimer->start(kGlobalTimeoutMs);
        emit transferProgress(data.size(), data.size());
    } else {
        // Download: data arrives via onDataReadyRead
        m_dataBuffer.clear();
    }
}

void FtpClient::onDataReadyRead()
{
    if (!m_isUpload) {
        m_dataBuffer += m_data->readAll();
        m_timeoutTimer->start(kGlobalTimeoutMs);
    }
}

void FtpClient::onDataFinished()
{
    if (!m_isUpload && !m_dataBuffer.isEmpty()) {
        QFile file(m_localPath);
        if (file.open(QIODevice::WriteOnly)) {
            file.write(m_dataBuffer);
            file.close();
        }
        m_lastProgressBytes = m_dataBuffer.size();
        m_timeoutTimer->stop();
        emit transferProgress(m_dataBuffer.size(), m_dataBuffer.size());
    }
    m_data->disconnect();
}

void FtpClient::onSocketError(QAbstractSocket::SocketError error)
{
    Q_UNUSED(error);
    abortTransfer(m_control->errorString());
}

void FtpClient::abortTransfer(const QString& error)
{
    m_state = State::Idle;
    m_replyBuffer.clear();
    m_dataBuffer.clear();
    m_timeoutTimer->stop();
    if (m_data->state() != QTcpSocket::UnconnectedState)
        m_data->abort();
    if (m_control->state() != QTcpSocket::UnconnectedState)
        m_control->abort();
    emit transferFailed(error);
}
