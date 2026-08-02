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

    void testWindows1252EscapedHex()
    {
        // \'e9 is 0xE9 = "é" in Windows-1252. Decoding the bytes as UTF-8
        // (the old behaviour) corrupted this.
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(
            QByteArrayLiteral("{\\rtf1\\ansi\\deff0 caf\\'e9}"), &doc);
        QVERIFY(ok);
        QVERIFY(doc.toPlainText().trimmed().contains(QStringLiteral("café")));
    }

    void testWindows1252LiteralHighByte()
    {
        // A raw 0xE9 byte in the body, as emitted by many ANSI RTF writers.
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 caf\xE9}");
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QVERIFY(doc.toPlainText().trimmed().contains(QStringLiteral("café")));
    }

    void testWindows1252SmartQuote()
    {
        // 0x93 = " in Windows-1252 (must not become a control char).
        QByteArray rtf = QByteArrayLiteral("{\\rtf1\\ansi\\deff0 \\'93quoted\\'94}");
        QTextDocument doc;
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);
        QVERIFY(doc.toPlainText().contains(QStringLiteral("\u201cquoted\u201d")));
    }

    void testColorTable()
    {
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral(
            "{\\rtf1\\ansi\\deff0 {\\colortbl;\\red0\\green0\\blue0;"
            "\\red255\\green0\\blue0;}\\cf2 red text}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);

        bool foundRed = false;
        for (QTextBlock block = doc.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                if (it.fragment().charFormat().foreground().color() == QColor(Qt::red))
                    foundRed = true;
            }
        }
        QVERIFY(foundRed);
    }

    void testColorTableDefaultIndex()
    {
        // \cf1 refers to the first color defined after the auto/black entry.
        QTextDocument doc;
        QByteArray rtf = QByteArrayLiteral(
            "{\\rtf1\\ansi\\deff0 {\\colortbl;\\red0\\green128\\blue0;}"
            "\\cf1 green text}");
        bool ok = RtfConverter::loadFromRtf(rtf, &doc);
        QVERIFY(ok);

        bool foundGreen = false;
        for (QTextBlock block = doc.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                if (it.fragment().charFormat().foreground().color() == QColor(0, 128, 0))
                    foundGreen = true;
            }
        }
        QVERIFY(foundGreen);
    }
};

QTEST_MAIN(TestRtfConverter)
#include "test_RtfConverter.moc"
