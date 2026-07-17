#include "NetworkUtils.h"

#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QNetworkRequest>
#include <QUrl>
#include <QVersionNumber>

static const QString s_defaultUpdateUrl = QStringLiteral("https://documenteditor.net/version.txt");

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

    QNetworkRequest request(url);
    request.setHeader(QNetworkRequest::UserAgentHeader,
                      QStringLiteral("Document.Editor Update Checker"));

    QNetworkReply *reply = m_nam->get(request);
    connect(reply, &QNetworkReply::finished, this, [this, reply, currentVersion]() {
        onVersionReply(reply, currentVersion);
    });
}

void NetworkUtils::onVersionReply(QNetworkReply *reply, const QString &currentVersion)
{
    reply->deleteLater();

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
        emit upToDate();
        return;
    }

    if (latest > current)
        emit updateAvailable(latestVersion,
            QStringLiteral("https://documenteditor.net/download/"));
    else
        emit upToDate();
}
