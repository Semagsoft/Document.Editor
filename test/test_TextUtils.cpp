#include <QTest>
#include "utils/TextUtils.h"

class TestTextUtils : public QObject
{
    Q_OBJECT

private slots:
    void testToUpperCase()
    {
        QCOMPARE(TextUtils::toUpperCase(QStringLiteral("hello")),   QStringLiteral("HELLO"));
        QCOMPARE(TextUtils::toUpperCase(QStringLiteral("Hello")),   QStringLiteral("HELLO"));
        QCOMPARE(TextUtils::toUpperCase(QStringLiteral("HELLO")),   QStringLiteral("HELLO"));
        QCOMPARE(TextUtils::toUpperCase(QString()),                  QString());
    }

    void testToLowerCase()
    {
        QCOMPARE(TextUtils::toLowerCase(QStringLiteral("HELLO")),   QStringLiteral("hello"));
        QCOMPARE(TextUtils::toLowerCase(QStringLiteral("Hello")),   QStringLiteral("hello"));
        QCOMPARE(TextUtils::toLowerCase(QString()),                  QString());
    }

    void testTitleCase()
    {
        QCOMPARE(TextUtils::titleCase(QStringLiteral("hello world")), QStringLiteral("Hello World"));
        QCOMPARE(TextUtils::titleCase(QStringLiteral("HELLO WORLD")), QStringLiteral("Hello World"));
        QCOMPARE(TextUtils::titleCase(QString()),                      QString());
    }

    void testSentenceCase()
    {
        QCOMPARE(TextUtils::sentenceCase(QStringLiteral("hello world")), QStringLiteral("Hello world"));
        QCOMPARE(TextUtils::sentenceCase(QStringLiteral("hello. world")), QStringLiteral("Hello. World"));
        QCOMPARE(TextUtils::sentenceCase(QString()),                      QString());
    }

    void testToggleCase()
    {
        QCOMPARE(TextUtils::toggleCase(QStringLiteral("Hello")),    QStringLiteral("hELLO"));
        QCOMPARE(TextUtils::toggleCase(QStringLiteral("HELLO")),    QStringLiteral("hello"));
        QCOMPARE(TextUtils::toggleCase(QStringLiteral("hello")),    QStringLiteral("HELLO"));
        QCOMPARE(TextUtils::toggleCase(QString()),                   QString());
    }

    void testWordCount()
    {
        QCOMPARE(TextUtils::wordCount(QStringLiteral("hello world")),           2);
        QCOMPARE(TextUtils::wordCount(QStringLiteral("hello")),                 1);
        QCOMPARE(TextUtils::wordCount(QString()),                               0);
        QCOMPARE(TextUtils::wordCount(QStringLiteral("")),                      0);
        QCOMPARE(TextUtils::wordCount(QStringLiteral("   ")),                   0);
        QCOMPARE(TextUtils::wordCount(QStringLiteral("hello   world")),         2);
    }

    void testCharacterCount()
    {
        QCOMPARE(TextUtils::characterCount(QStringLiteral("hello")),                 5);
        QCOMPARE(TextUtils::characterCount(QStringLiteral("hello world")),           11);
        QCOMPARE(TextUtils::characterCount(QStringLiteral("hello world"), false),    10);
        QCOMPARE(TextUtils::characterCount(QString()),                               0);
    }
};

QTEST_MAIN(TestTextUtils)
#include "test_TextUtils.moc"
