#include <QTest>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextImageFormat>
#include <QImage>
#include <QMarginsF>
#include <QColor>
#include "converters/XamlConverter.h"

class TestXamlConverter : public QObject
{
    Q_OBJECT

private slots:
    void testSaveAndLoadRoundTrip()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        doc.setPlainText(QStringLiteral("Hello World"));

        QMarginsF margins(72, 72, 72, 72);
        QColor bg(Qt::white);

        QString xaml = XamlConverter::saveToXaml(&doc, margins);

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = XamlConverter::loadFromXaml(xaml, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);
        QCOMPARE(doc2.toPlainText().trimmed(), QStringLiteral("Hello World"));
        QCOMPARE(loadedMargins.left(), 72.0);
        QCOMPARE(loadedMargins.top(), 72.0);
        QCOMPARE(doc2.pageSize().width(), 816.0);
        QCOMPARE(doc2.pageSize().height(), 1056.0);
    }

    void testEmptyDocument()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(612, 792));
        QMarginsF margins;
        QString xaml = XamlConverter::saveToXaml(&doc, margins);

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = XamlConverter::loadFromXaml(xaml, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);
        QCOMPARE(doc2.pageSize().width(), 612.0);
        QCOMPARE(doc2.pageSize().height(), 792.0);
    }

    void testInvalidXml()
    {
        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        bool ok = XamlConverter::loadFromXaml(QStringLiteral("not xml"), &doc, margins, bg);
        QVERIFY(!ok);
    }

    void testBoldFormatting()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        QTextCharFormat fmt;
        fmt.setFontWeight(QFont::Bold);
        cursor.insertText(QStringLiteral("Bold text"), fmt);

        QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());
        QVERIFY(xaml.contains(QStringLiteral("Bold=\"True\"")));
    }

    void testSoftLineBreakRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        cursor.insertText(QStringLiteral("one"));
        cursor.insertText(QString(QChar::LineSeparator));
        cursor.insertText(QStringLiteral("two"));

        QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());
        QVERIFY(xaml.contains(QStringLiteral("LineBreak")));

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = XamlConverter::loadFromXaml(xaml, &doc2, margins, bg);
        QVERIFY(ok);
        QCOMPARE(doc2.blockCount(), 1); // a soft break must not split the paragraph
        QVERIFY(doc2.begin().text().contains(QString(QChar::LineSeparator)));
    }

    void testFontFamilyDoesNotClobberBold()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        QTextCharFormat fmt;
        fmt.setFontWeight(QFont::Bold);
        fmt.setFontItalic(true);
        fmt.setFontFamilies({ QStringLiteral("Comic Sans MS") });
        cursor.insertText(QStringLiteral("styled text"), fmt);

        QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());
        QVERIFY(xaml.contains(QStringLiteral("Bold=\"True\"")));
        QVERIFY(xaml.contains(QStringLiteral("Italic=\"True\"")));
        QVERIFY(xaml.contains(QStringLiteral("Comic Sans MS")));

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = XamlConverter::loadFromXaml(xaml, &doc2, margins, bg);
        QVERIFY(ok);

        QTextCharFormat loaded;
        for (QTextBlock block = doc2.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                QTextFragment f = it.fragment();
                if (f.text().contains(QStringLiteral("styled"))) {
                    loaded = f.charFormat();
                    break;
                }
            }
        }
        QCOMPARE(loaded.fontWeight(), QFont::Bold);
        QVERIFY(loaded.fontItalic());
        QVERIFY(loaded.fontFamilies().toStringList().contains(QStringLiteral("Comic Sans MS")));
    }

    void testMultipleImagesSameParagraph()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QImage img1(16, 16, QImage::Format_ARGB32);
        img1.fill(Qt::blue);
        QTextImageFormat imgFmt1;
        QString name1 = QStringLiteral("xaml_img_a.png");
        doc.addResource(QTextDocument::ImageResource, QUrl(name1), img1);
        imgFmt1.setName(name1);
        imgFmt1.setWidth(32);
        imgFmt1.setHeight(32);

        QImage img2(16, 16, QImage::Format_ARGB32);
        img2.fill(Qt::red);
        QTextImageFormat imgFmt2;
        QString name2 = QStringLiteral("xaml_img_b.png");
        doc.addResource(QTextDocument::ImageResource, QUrl(name2), img2);
        imgFmt2.setName(name2);
        imgFmt2.setWidth(32);
        imgFmt2.setHeight(32);

        cursor.insertImage(imgFmt1);
        cursor.insertImage(imgFmt2);

        QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());
        QVERIFY(!xaml.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = XamlConverter::loadFromXaml(xaml, &doc2, margins, bg);
        QVERIFY(ok);

        int imageCount = 0;
        for (QTextBlock block = doc2.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                if (it.fragment().charFormat().isImageFormat())
                    ++imageCount;
            }
        }
        QCOMPARE(imageCount, 2);
    }

    void testListRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        QTextListFormat lf;
        lf.setStyle(QTextListFormat::ListDisc);
        cursor.insertText(QStringLiteral("Item A"));
        cursor.insertBlock();
        cursor.createList(lf);
        cursor.insertText(QStringLiteral("Item B"));
        cursor.insertBlock();
        cursor.insertText(QStringLiteral("Item C"));

        QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());
        QVERIFY(!xaml.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = XamlConverter::loadFromXaml(xaml, &doc2, margins, bg);
        QVERIFY(ok);
        QCOMPARE(doc2.toPlainText(), doc.toPlainText());

        int srcInList = 0, dstInList = 0;
        for (QTextBlock b = doc.begin(); b.isValid(); b = b.next())
            if (b.textList()) ++srcInList;
        for (QTextBlock b = doc2.begin(); b.isValid(); b = b.next())
            if (b.textList()) ++dstInList;
        QCOMPARE(dstInList, srcInList);
    }

    void testTableRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        cursor.insertText(QStringLiteral("Before table"));

        QTextTable *table = cursor.insertTable(2, 2);

        QTextTableCell cell00 = table->cellAt(0, 0);
        QTextCursor c00 = cell00.firstCursorPosition();
        c00.insertText(QStringLiteral("First para"));
        c00.insertBlock();
        c00.insertText(QStringLiteral("Second para"));

        QTextTableCell cell01 = table->cellAt(0, 1);
        QTextCursor c01 = cell01.firstCursorPosition();
        c01.insertText(QStringLiteral("Cell A"));

        QTextTableCell cell10 = table->cellAt(1, 0);
        QTextCursor c10 = cell10.firstCursorPosition();
        c10.insertText(QStringLiteral("Cell B"));

        QTextTableCell cell11 = table->cellAt(1, 1);
        QTextCursor c11 = cell11.firstCursorPosition();
        c11.insertText(QStringLiteral("Cell C"));

        cursor.movePosition(QTextCursor::End);
        cursor.insertText(QStringLiteral("After table"));

        QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());
        QVERIFY(!xaml.isEmpty());
        QVERIFY(xaml.contains(QStringLiteral("<Table")));

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = XamlConverter::loadFromXaml(xaml, &doc2, margins, bg);
        QVERIFY(ok);

        // Text before and after the table must be preserved.
        QString text = doc2.toPlainText();
        QVERIFY(text.contains(QStringLiteral("Before table")));
        QVERIFY(text.contains(QStringLiteral("After table")));

        // The table structure must survive the round trip.
        QTextTable *loaded = nullptr;
        for (QTextFrame *frame = doc2.rootFrame(); frame; frame = frame->parentFrame()) {
            for (QTextFrame::iterator it = frame->begin(); !it.atEnd(); ++it) {
                if (QTextFrame *child = it.currentFrame()) {
                    if (QTextTable *t = qobject_cast<QTextTable *>(child)) {
                        loaded = t;
                        break;
                    }
                }
            }
            if (loaded)
                break;
        }
        QVERIFY(loaded);
        QCOMPARE(loaded->rows(), 2);
        QCOMPARE(loaded->columns(), 2);

        // Cell content must be intact (multi-paragraph cell included).
        QTextCursor l00 = loaded->cellAt(0, 0).firstCursorPosition();
        QTextBlock block = l00.block();
        QCOMPARE(block.text(), QStringLiteral("First para"));
        block = block.next();
        QVERIFY(block.isValid());
        QCOMPARE(block.text(), QStringLiteral("Second para"));
        QCOMPARE(loaded->cellAt(0, 1).firstCursorPosition().block().text(),
                 QStringLiteral("Cell A"));
        QCOMPARE(loaded->cellAt(1, 0).firstCursorPosition().block().text(),
                 QStringLiteral("Cell B"));
        QCOMPARE(loaded->cellAt(1, 1).firstCursorPosition().block().text(),
                 QStringLiteral("Cell C"));
    }
};

QTEST_MAIN(TestXamlConverter)
#include "test_XamlConverter.moc"
