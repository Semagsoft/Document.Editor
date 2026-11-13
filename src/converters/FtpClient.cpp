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
    connect(m_data, &QTcpSocket::errorOccurred, this, &FtpClient::onDataError);

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

bool FtpClient::parseUrl(const QString& url, QString& host, quint16& port,
    QString& path, QString& error) const
{
    auto hasCrlf = [](const QString& in) {
        return in.contains(QLatin1Char('\r')) || in.contains(QLatin1Char('\n'));
    };
    QUrl qurl(url);
    if (!qurl.isValid()) {
        error = tr("The URL is not valid.");
        return false;
    }
    if (hasCrlf(url)) {
        error = tr("The URL contains invalid characters.");
        return false;
    }
    host = qurl.host();
    if (host.isEmpty() || hasCrlf(host)) {
        error = tr("The URL does not contain a valid host.");
        return false;
    }
    const int p = qurl.port(21);
    if (p < 1 || p > 65535) {
        error = tr("The port %1 is out of range.").arg(p);
        return false;
    }
    port = static_cast<quint16>(p);
    path = qurl.path();
    if (hasCrlf(path)) {
        error = tr("The remote path contains invalid characters.");
        return false;
    }
    if (path.isEmpty())
        path = QStringLiteral("/");
    return true;
}

void FtpClient::upload(const QString& localPath, const QString& remoteUrl,
    const QString& username, const QString& password)
{
    if (isBusy())
        return;

    m_localPath = localPath;
    QString urlError;
    if (!parseUrl(remoteUrl, m_host, m_port, m_remotePath, urlError)) {
        m_state = State::Idle;
        emit transferFailed(urlError);
        return;
    }
    if (username.contains(QLatin1Char('\r')) || username.contains(QLatin1Char('\n'))
        || password.contains(QLatin1Char('\r')) || password.contains(QLatin1Char('\n'))) {
        m_state = State::Idle;
        emit transferFailed(tr("The credentials contain invalid characters."));
        return;
    }
    m_username = username.isEmpty() ? QStringLiteral("anonymous") : username;
    m_password = password.isEmpty() ? QStringLiteral("anonymous@") : password;
    m_isUpload = true;
    m_dataTransferDone = false;
    m_controlTransferDone = false;
    m_finalized = false;

    // Reset sockets from any previous session: connectToHost() on an already
    // connected socket does not reliably start a fresh FTP session.
    m_replyBuffer.clear();
    m_control->abort();
    m_data->abort();

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
    QString urlError;
    if (!parseUrl(remoteUrl, m_host, m_port, m_remotePath, urlError)) {
        m_state = State::Idle;
        emit transferFailed(urlError);
        return;
    }
    if (username.contains(QLatin1Char('\r')) || username.contains(QLatin1Char('\n'))
        || password.contains(QLatin1Char('\r')) || password.contains(QLatin1Char('\n'))) {
        m_state = State::Idle;
        emit transferFailed(tr("The credentials contain invalid characters."));
        return;
    }
    m_username = username.isEmpty() ? QStringLiteral("anonymous") : username;
    m_password = password.isEmpty() ? QStringLiteral("anonymous@") : password;
    m_isUpload = false;
    m_dataTransferDone = false;
    m_controlTransferDone = false;
    m_finalized = false;

    // Reset sockets from any previous session: connectToHost() on an already
    // connected socket does not reliably start a fresh FTP session.
    m_replyBuffer.clear();
    m_control->abort();
    m_data->abort();

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
    if (m_control->state() == QTcpSocket::ConnectedState) {
        sendCommand(QStringLiteral("QUIT"));
        m_control->disconnectFromHost();
    } else if (m_control->state() != QTcpSocket::UnconnectedState) {
        m_control->abort();
    }
    m_state = State::Idle;
    m_replyBuffer.clear();
    m_dataBuffer.clear();
    m_dataTransferDone = false;
    m_controlTransferDone = false;
    m_finalized = false;
    m_timeoutTimer->stop();
}

void FtpClient::sendCommand(const QString& cmd)
{
    if (cmd.contains(QLatin1Char('\r')) || cmd.contains(QLatin1Char('\n'))) {
        // Never let user input smuggle extra FTP commands into the control channel.
        abortTransfer(tr("Invalid characters in FTP command."));
        return;
    }
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
            if (code == 220) {
                m_state = State::Connected;
                sendCommand(QStringLiteral("USER %1").arg(m_username));
            } else {
                abortTransfer(tr("Unexpected FTP response: %1").arg(QString::fromUtf8(line)));
            }
            break;

        case State::Idle:
            // Accept an unsolicited greeting, and ignore benign 2xx replies
            // (e.g. a late 226) that follow a completed/cancelled transfer.
            if (code == 220) {
                m_state = State::Connected;
                sendCommand(QStringLiteral("USER %1").arg(m_username));
            } else if (code >= 200 && code < 300) {
                // no-op
            } else {
                abortTransfer(tr("Unexpected FTP response: %1").arg(QString::fromUtf8(line)));
            }
            break;

        case State::Connected:
            if (code == 331)
                sendCommand(QStringLiteral("PASS %1").arg(m_password));
            else if (code == 230) {
                m_state = State::LoggedIn;
                sendCommand(QStringLiteral("PASV"));
            } else {
                abortTransfer(tr("Login failed: %1").arg(QString::fromUtf8(line)));
            }
            break;

        case State::LoggedIn:
            if (code == 227) {
                // Parse PASV response: 227 Entering Passive Mode (h1,h2,h3,h4,p1,p2)
                QRegularExpression re(QStringLiteral(
                    "\\((\\d{1,3}),(\\d{1,3}),(\\d{1,3}),(\\d{1,3}),(\\d{1,3}),(\\d{1,3})\\)"));
                auto match = re.match(QString::fromUtf8(line));
                bool pasvOk = match.hasMatch();
                quint16 pasvPort = 0;
                if (pasvOk) {
                    for (int gi = 1; gi <= 6; ++gi) {
                        if (match.captured(gi).toInt() > 255)
                            pasvOk = false;
                    }
                    int dataPort = match.captured(5).toInt() * 256 + match.captured(6).toInt();
                    if (dataPort < 1 || dataPort > 65535)
                        pasvOk = false;
                    else
                        pasvPort = static_cast<quint16>(dataPort);
                }
                if (pasvOk) {
                    // RFC 2577: do not trust the PASV-provided IP; reuse the
                    // control connection's host to avoid FTP bounce/SSRF.
                    m_dataHost = m_host;
                    m_dataPort = pasvPort;
                    m_state = State::Passive;
                    connectDataChannel();
                } else {
                    abortTransfer(tr("Failed to parse PASV response"));
                }
            } else {
                abortTransfer(tr("Failed to enter passive mode: %1").arg(QString::fromUtf8(line)));
            }
            break;

        case State::Passive:
            // 1xx = data connection open (e.g. 150), begin transfer now
            if (code >= 100 && code < 200) {
                m_state = State::Transferring;
                startTransfer();
            } else if (code >= 200 && code < 300) {
                // Some servers omit the 1xx preliminary reply
                finishTransfer();
            } else {
                abortTransfer(tr("Transfer rejected: %1").arg(QString::fromUtf8(line)));
            }
            break;

        case State::Transferring:
            // 226 = transfer complete
            if (code >= 200 && code < 300) {
                finishTransfer();
            } else {
                abortTransfer(tr("Transfer failed: %1").arg(QString::fromUtf8(line)));
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
    m_timeoutTimer->start(kGlobalTimeoutMs);
    if (m_isUpload) {
        QFile file(m_localPath);
        if (!file.open(QIODevice::ReadOnly)) {
            abortTransfer(tr("Cannot open local file: %1").arg(m_localPath));
            return;
        }
        QByteArray data = file.readAll();
        file.close();
        // Write the payload, then half-close the data channel. QAbstractSocket
        // flushes any pending data before the connection actually closes, so
        // the server sees a clean EOF and replies 226.
        m_data->write(data);
        m_data->disconnectFromHost();
        m_lastProgressBytes = data.size();
        emit transferProgress(data.size(), data.size());
    } else {
        // Download: data arrives via onDataReadyRead until the server closes
        // the data channel (onDataFinished) or the 226 reply is processed.
        m_dataBuffer.clear();
    }
}

void FtpClient::finishTransfer()
{
    m_timeoutTimer->stop();
    if (m_isUpload) {
        m_state = State::Idle;
        emit uploadCompleted();
    } else {
        m_controlTransferDone = true;
        tryFinalizeDownload();
    }
}

bool FtpClient::writeDownloadedData(QString& error)
{
    QFile file(m_localPath);
    if (!file.open(QIODevice::WriteOnly)) {
        error = tr("Could not write the file: %1").arg(file.errorString());
        return false;
    }
    const qint64 written = file.write(m_dataBuffer);
    file.close();
    if (written != m_dataBuffer.size()) {
        error = tr("Could not write the file: short write (%1 of %2 bytes)")
                    .arg(written).arg(m_dataBuffer.size());
        return false;
    }
    return true;
}

void FtpClient::tryFinalizeDownload()
{
    if (m_finalized || m_isUpload)
        return;
    // Both sides must agree: the data socket reached EOF and the control
    // channel reported success (226/2xx).
    if (!m_dataTransferDone || !m_controlTransferDone)
        return;
    m_finalized = true;
    QString error;
    if (!writeDownloadedData(error)) {
        m_state = State::Idle;
        m_dataBuffer.clear();
        emit transferFailed(error);
        return;
    }
    m_lastProgressBytes = m_dataBuffer.size();
    emit transferProgress(m_dataBuffer.size(), m_dataBuffer.size());
    m_dataBuffer.clear();
    m_state = State::Idle;
    emit downloadCompleted();
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
    if (m_isUpload)
        return;
    m_dataTransferDone = true;
    tryFinalizeDownload();
}

void FtpClient::onSocketError(QAbstractSocket::SocketError error)
{
    Q_UNUSED(error);
    abortTransfer(m_control->errorString());
}

void FtpClient::onDataError(QAbstractSocket::SocketError error)
{
    // RemoteHostClosedError is the normal server-side EOF at the end of a
    // download; it is completed by onDataFinished()/the 226 reply, not a failure.
    if (error == QAbstractSocket::RemoteHostClosedError)
        return;
    if (m_state == State::Passive || m_state == State::Transferring)
        abortTransfer(m_data->errorString());
}

void FtpClient::abortTransfer(const QString& error)
{
    m_state = State::Idle;
    m_replyBuffer.clear();
    m_dataBuffer.clear();
    m_dataTransferDone = false;
    m_controlTransferDone = false;
    m_finalized = false;
    m_timeoutTimer->stop();
    if (m_data->state() != QTcpSocket::UnconnectedState)
        m_data->abort();
    if (m_control->state() != QTcpSocket::UnconnectedState)
        m_control->abort();
    emit transferFailed(error);
}
