#include <QTest>
#include "utils/NetworkUtils.h"

class TestNetworkUtils : public QObject
{
    Q_OBJECT

private slots:
    void testDefaultUpdateUrl()
    {
        QString url = NetworkUtils::defaultUpdateUrl();
        QVERIFY(!url.isEmpty());
        QVERIFY(url.startsWith(QStringLiteral("https://")));
    }

    void testCustomUpdateUrl()
    {
        NetworkUtils netUtils;
        QCOMPARE(netUtils.updateUrl(), NetworkUtils::defaultUpdateUrl());

        netUtils.setUpdateUrl(QStringLiteral("https://example.com/version.txt"));
        QCOMPARE(netUtils.updateUrl(), QStringLiteral("https://example.com/version.txt"));
    }
};

QTEST_MAIN(TestNetworkUtils)
#include "test_NetworkUtils.moc"
