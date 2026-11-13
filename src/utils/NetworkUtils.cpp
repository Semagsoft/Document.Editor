#include "NetworkUtils.h"

#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QNetworkRequest>
#include <QTimer>
#include <QUrl>
#include <QVersionNumber>

static const QString s_defaultUpdateUrl = QStringLiteral("https://documenteditor.net/version.txt");
constexpr int kUpdateTimeoutMs = 15000;

NetworkUtils::NetworkUtils(QObject *parent)
    : QObject(parent)
    , m_nam(new QNetworkAccessManager(this))
{
}

QString NetworkUtils::defaultUpdateUrl()
{
    return s_defaultUpdateUrl;
}

void NetworkUtils::setUpdateUrl(const QString &url)
{
    m_updateUrl = url;
}

QString NetworkUtils::updateUrl() const
{
    return m_updateUrl.isEmpty() ? s_defaultUpdateUrl : m_updateUrl;
}

void NetworkUtils::checkForUpdates(const QString &currentVersion,
                                   const QString &updateUrl)
{
    QString url = updateUrl.isEmpty() ? this->updateUrl() : updateUrl;
    if (url.isEmpty())
        url = s_defaultUpdateUrl;

    // Refuse plaintext HTTP: an update list received in the clear can be
    // tampered with to push a malicious "new version".
    if (QUrl(url).scheme().compare(QLatin1String("https"), Qt::CaseInsensitive) != 0) {
        emit updateCheckError(tr("The update URL must use HTTPS"));
        return;
    }

    QNetworkRequest request(url);
    request.setHeader(QNetworkRequest::UserAgentHeader,
                      QStringLiteral("Document.Editor Update Checker"));

    QNetworkReply *reply = m_nam->get(request);

    // Bound the reply: absolute size cap and a hard timeout. An untrusted
    // server must not stream unbounded data or hang the check forever.
    constexpr qint64 kMaxVersionBytes = 64 * 1024;
    connect(reply, &QNetworkReply::readyRead, this, [reply]() {
        if (reply->bytesAvailable() > kMaxVersionBytes) {
            reply->setProperty("oversized", true);
            reply->abort();
        }
    });

    auto timeout = new QTimer(this);
    timeout->setSingleShot(true);
    connect(timeout, &QTimer::timeout, this, [this, reply, timeout]() {
        if (reply->isRunning()) {
            reply->setProperty("timedOut", true);
            reply->abort();
        }
        timeout->deleteLater();
    });
    connect(reply, &QNetworkReply::finished, timeout, &QTimer::deleteLater);
    timeout->start(kUpdateTimeoutMs);

    connect(reply, &QNetworkReply::finished, this, [this, reply, currentVersion]() {
        onVersionReply(reply, currentVersion);
    });
}

void NetworkUtils::onVersionReply(QNetworkReply *reply, const QString &currentVersion)
{
    reply->deleteLater();

    if (reply->property("timedOut").toBool()) {
        emit updateCheckError(tr("Update check timed out"));
        return;
    }
    if (reply->property("oversized").toBool()) {
        emit updateCheckError(tr("Update server response was too large"));
        return;
    }
    if (reply->error() != QNetworkReply::NoError) {
        emit updateCheckError(reply->errorString());
        return;
    }

    QString latestVersion = QString::fromUtf8(reply->readAll()).trimmed();
    if (latestVersion.isEmpty()) {
        emit updateCheckError(tr("Empty response from update server"));
        return;
    }

    QVersionNumber current = QVersionNumber::fromString(currentVersion);
    QVersionNumber latest = QVersionNumber::fromString(latestVersion);

    if (current.isNull() || latest.isNull()) {
        // A response we can't parse is an error, not "up to date" — silently
        // claiming up-to-date hides broken/malicious update endpoints.
        emit updateCheckError(tr("Could not parse the version returned by the update server"));
        return;
    }

    if (latest > current)
        emit updateAvailable(latestVersion,
            QStringLiteral("https://documenteditor.net/download/"));
    else
        emit upToDate();
}
