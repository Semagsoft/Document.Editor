#include <QTest>
#include "utils/SpellChecker.h"

class TestSpellChecker : public QObject
{
    Q_OBJECT

private slots:
    void testCleanWord()
    {
        SpellChecker sc;
        // Can't test cleanWord directly since it's private;
        // test via isMisspelled which will return false for empty dict
        QVERIFY(!sc.hasDictionary());
        QVERIFY(!sc.isMisspelled(QStringLiteral("hello")));
    }

    void testEmptySuggestions()
    {
        SpellChecker sc;
        QVERIFY(!sc.hasDictionary());
        QVERIFY(sc.suggestions(QStringLiteral("hello")).isEmpty());
    }

    void testAddToUserDictionary()
    {
        SpellChecker sc;
        sc.addToUserDictionary(QStringLiteral("testword"));
        QStringList userDict = sc.userDictionary();
        QVERIFY(userDict.contains(QStringLiteral("testword")));

        sc.addToUserDictionary(QStringLiteral("ANOTHER"));
        userDict = sc.userDictionary();
        QVERIFY(userDict.contains(QStringLiteral("another")));
    }

    void testEmptyTextFindMisspelled()
    {
        SpellChecker sc;
        QStringList result = sc.findMisspelled(QString());
        QVERIFY(result.isEmpty());

        result = sc.findMisspelled(QStringLiteral("   "));
        QVERIFY(result.isEmpty());
    }
};

QTEST_MAIN(TestSpellChecker)
#include "test_SpellChecker.moc"
