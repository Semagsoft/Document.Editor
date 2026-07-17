#include <QTest>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextCharFormat>
#include <QTextBlock>
#include <QFont>
#include "converters/RtfConverter.h"

class TestRtfConverter : public QObject
{
    Q_OBJECT

private slots:
    void testInvalidRtf()
    {
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(QByteArrayLiteral("not rtf"), &doc);
        QVERIFY(!ok);
    }

    void testEmptyRtf()
    {
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(
            QByteArrayLiteral("{\\rtf1\\ansi\\deff0}"), &doc);
        QVERIFY(ok);
    }

    void testPlainText()
    {
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(
            QByteArrayLiteral("{\\rtf1\\ansi\\deff0 Hello World}"), &doc);
        QVERIFY(ok);
        QCOMPARE(doc.toPlainText().trimmed(), QStringLiteral("Hello World"));
    }

    void testBoldFormatting()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\b Bold \\b0 normal}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QString plain = doc.toPlainText().trimmed();
        QVERIFY(plain.contains(QStringLiteral("Bold")));
    }

    void testItalicFormatting()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\i Italic \\i0 normal}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QString plain = doc.toPlainText().trimmed();
        QVERIFY(plain.contains(QStringLiteral("Italic")));
    }

    void testUnderlineFormatting()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\ul Underlined \\ulnone normal}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QString plain = doc.toPlainText().trimmed();
        QVERIFY(plain.contains(QStringLiteral("Underlined")));
    }

    void testBoldItalicCombined()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\b \\i BoldItalic}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
    }

    void testParagraphBreak()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 First\\par Second}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QCOMPARE(doc.blockCount(), 2);
    }

    void testUnicode()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\u8364?}"); // euro sign
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
    }

    void testFontSize()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\fs48 Large text}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QString plain = doc.toPlainText().trimmed();
        QVERIFY(plain.contains(QStringLiteral("Large")));
    }

    void testSaveAndLoadRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        QTextCharFormat fmt;
        fmt.setFontWeight(QFont::Bold);
        cursor.insertText(QStringLiteral("Bold text"), fmt);

        QByteArray rtf = RtfConverter::saveToRtf(&doc);
        QVERIFY(!rtf.isEmpty());
        QVERIFY(rtf.contains("\\b "));

        QTextDocument doc2;
        bool ok = RtfConverter::loadFromRtf(rtf, &doc2);
        QVERIFY(ok);
        QCOMPARE(doc2.toPlainText().trimmed(), QStringLiteral("Bold text"));
    }

    void testCenterAlignment()
    {
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(
            QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\qc Centered text\\par}"), &doc);
        QVERIFY(ok);
        QCOMPARE(doc.firstBlock().blockFormat().alignment(), Qt::AlignCenter);
    }

    void testRightAlignment()
    {
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(
            QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\qr Right text\\par}"), &doc);
        QVERIFY(ok);
        QCOMPARE(doc.firstBlock().blockFormat().alignment(), Qt::AlignRight);
    }

    void testScopedFormatting()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral(
            "{\\rtf1\\ansi\\deff0 Normal {\\b Bold inside} normal again}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QString plain = doc.toPlainText().trimmed();
        QVERIFY(plain.contains(QStringLiteral("Normal")));
        QVERIFY(plain.contains(QStringLiteral("Bold inside")));
        QVERIFY(plain.contains(QStringLiteral("normal again")));
    }
};

QTEST_MAIN(TestRtfConverter)
#include "test_RtfConverter.moc"
