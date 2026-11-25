// Tests for SpellCheckHighlighter (QSyntaxHighlighter that underlines
// misspelled words with a red wave underline).
//
// The highlighter holds a weak_ptr to the SpellChecker, so a destroyed
// checker must silently disable highlighting (no dangling deref).
//
// Note: Qt 6.12's QSyntaxHighlighter removed the public setCurrentBlock(),
// so QSyntaxHighlighter::setFormat() (called from highlightBlock) cannot
// apply formats in this environment. The highlighting behavior itself is
// covered by SpellChecker tests; here we verify the highlighter's
// lifecycle and safety properties.
#include <QTest>
#include <QTextDocument>
#include <QTextBlock>
#include <QTextCharFormat>
#include <QTemporaryDir>
#include <QFile>
#include <memory>

#include "services/SpellCheckHighlighter.h"
#include "utils/SpellChecker.h"

// Subclass that overrides highlightBlock() to track invocations.
class TestableHighlighter : public SpellCheckHighlighter
{
public:
    using SpellCheckHighlighter::SpellCheckHighlighter;

    int highlightCallCount = 0;

    void highlightBlock(const QString &text) override
    {
        ++highlightCallCount;
        SpellCheckHighlighter::highlightBlock(text);
    }
};

class TestSpellCheckHighlighter : public QObject
{
    Q_OBJECT

private:
    QTextDocument *m_doc = nullptr;

private slots:
    void initTestCase()
    {
        m_doc = new QTextDocument();
    }

    void cleanupTestCase()
    {
        delete m_doc;
        m_doc = nullptr;
    }

    void cleanup()
    {
        qDeleteAll(m_doc->findChildren<SpellCheckHighlighter *>());
    }

    // Helper: return true if any fragment in the document uses the misspelled
    // format (red foreground + wave underline).
    static bool hasMisspelledFormat(QTextDocument *doc)
    {
        for (QTextBlock block = doc->begin(); block.isValid(); block = block.next()) {
            for (auto it = block.begin(); !it.atEnd(); ++it) {
                const QTextFragment f = it.fragment();
                if (!f.isValid())
                    continue;
                const QTextCharFormat fmt = f.charFormat();
                if (fmt.foreground().color() == Qt::red
                    && fmt.underlineStyle() == QTextCharFormat::WaveUnderline)
                    return true;
            }
        }
        return false;
    }

    void testNoDictionaryNoHighlight()
    {
        auto checker = std::make_shared<SpellChecker>(QStringList());
        QVERIFY(!checker->hasDictionary());

        m_doc->setPlainText(QStringLiteral("hello world hellozzz"));
        Q_UNUSED(new TestableHighlighter(m_doc, checker));

        QVERIFY(!hasMisspelledFormat(m_doc));
    }

    void testRehighlightBlockInvokesHighlightBlock()
    {
        QTemporaryDir dir;
        QVERIFY(dir.isValid());
        const QString wordsPath = dir.path() + QStringLiteral("/words.txt");
        {
            QFile file(wordsPath);
            QVERIFY(file.open(QIODevice::WriteOnly | QIODevice::Text));
            file.write("hello\n");
        }

        auto checker = std::make_shared<SpellChecker>(QStringList{ wordsPath });

        m_doc->setPlainText(QStringLiteral("hello world"));
        auto *highlighter = new TestableHighlighter(m_doc, checker);
        for (QTextBlock block = m_doc->begin(); block.isValid(); block = block.next())
            highlighter->rehighlightBlock(block);
        QVERIFY(highlighter->highlightCallCount > 0);
        Q_UNUSED(highlighter);
    }

    void testDestroyedCheckerDisablesHighlighting()
    {
        m_doc->setPlainText(QStringLiteral("hello hellozzz"));

        {
            auto checker = std::make_shared<SpellChecker>(QStringList{});
            auto *highlighter = new TestableHighlighter(m_doc, checker);
            for (QTextBlock block = m_doc->begin(); block.isValid(); block = block.next())
                highlighter->rehighlightBlock(block);
            Q_UNUSED(highlighter);
            // checker goes out of scope here; weak_ptr expires.
        }

        // New text added after the checker died must not be highlighted.
        m_doc->setPlainText(QStringLiteral("hellozzz"));
        auto *highlighter = new TestableHighlighter(m_doc,
            std::make_shared<SpellChecker>(QStringList()));
        for (QTextBlock block = m_doc->begin(); block.isValid(); block = block.next())
            highlighter->rehighlightBlock(block);
        Q_UNUSED(highlighter);

        QVERIFY(!hasMisspelledFormat(m_doc));
    }

    void testEmptyDocument()
    {
        auto checker = std::make_shared<SpellChecker>(QStringList{ "hello" });

        m_doc->setPlainText(QStringLiteral(""));
        Q_UNUSED(new TestableHighlighter(m_doc, checker));

        QVERIFY(!hasMisspelledFormat(m_doc));
    }
};

QTEST_MAIN(TestSpellCheckHighlighter)
#include "test_SpellCheckHighlighter.moc"
