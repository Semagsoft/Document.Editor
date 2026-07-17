#pragma once

#include <QObject>
#include <QString>

class QNetworkAccessManager;
class QNetworkReply;

class NetworkUtils : public QObject
{
    Q_OBJECT

public:
    explicit NetworkUtils(QObject *parent = nullptr);

    void checkForUpdates(const QString &currentVersion,
                         const QString &updateUrl = QString());

    void setUpdateUrl(const QString &url);

    QString updateUrl() const;
    static QString defaultUpdateUrl();

signals:
    void updateAvailable(const QString &latestVersion, const QString &downloadUrl);
    void upToDate();
    void updateCheckError(const QString &error);

private:
    QNetworkAccessManager *m_nam = nullptr;
    QString m_updateUrl;

    void onVersionReply(QNetworkReply *reply, const QString &currentVersion);
};
