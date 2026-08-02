#include <QTest>
#include <QTemporaryDir>
#include <QFile>
#include "utils/SpellChecker.h"

class TestSpellChecker : public QObject
{
    Q_OBJECT

private slots:
    void testNoDictionary()
    {
        SpellChecker sc{QStringList()};
        QVERIFY(!sc.hasDictionary());
        QVERIFY(sc.suggestions(QStringLiteral("hello")).isEmpty());
        QVERIFY(!sc.isMisspelled(QStringLiteral("hello")));
    }

    void testWithDictionary()
    {
        QTemporaryDir dir;
        QVERIFY(dir.isValid());
        QString wordsPath = dir.path() + QStringLiteral("/words.txt");
        {
            QFile file(wordsPath);
            QVERIFY(file.open(QIODevice::WriteOnly | QIODevice::Text));
            file.write("hello\nworld\nquick\nbrown\n");
        }

        SpellChecker sc(QStringList{ wordsPath },
                        dir.path() + QStringLiteral("/user.txt"));
        QVERIFY(sc.hasDictionary());
        QVERIFY(!sc.isMisspelled(QStringLiteral("hello")));
        QVERIFY(sc.isMisspelled(QStringLiteral("hellozzz")));
        QVERIFY(sc.suggestions(QStringLiteral("helllo"))
                    .contains(QStringLiteral("hello")));
    }

    void testAddToUserDictionary()
    {
        QTemporaryDir dir;
        QVERIFY(dir.isValid());
        SpellChecker sc{QStringList(), dir.path() + QStringLiteral("/user.txt")};

        sc.addToUserDictionary(QStringLiteral("testword"));
        QVERIFY(sc.userDictionary().contains(QStringLiteral("testword")));

        sc.addToUserDictionary(QStringLiteral("ANOTHER"));
        QVERIFY(sc.userDictionary().contains(QStringLiteral("another")));
    }

    void testEmptyTextFindMisspelled()
    {
        SpellChecker sc{QStringList()};
        QStringList result = sc.findMisspelled(QString());
        QVERIFY(result.isEmpty());

        result = sc.findMisspelled(QStringLiteral("   "));
        QVERIFY(result.isEmpty());
    }
};

QTEST_MAIN(TestSpellChecker)
#include "test_SpellChecker.moc"
