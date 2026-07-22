#include <QTest>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextCharFormat>
#include <QTextBlockFormat>
#include <QTextList>
#include <QTextListFormat>
#include <QTextTable>
#include <QTextFrame>
#include <QMarginsF>
#include <QColor>
#include <QBuffer>
#include <QImage>

#include "converters/DocxConverter.h"

class TestDocxConverter : public QObject
{
    Q_OBJECT

private slots:
    void testSaveAndLoadRoundTrip()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        QTextCursor cursor(&doc);

        QTextCharFormat boldFmt;
        boldFmt.setFontWeight(QFont::Bold);
        boldFmt.setFontPointSize(16);
        cursor.insertText(QStringLiteral("Bold Heading"), boldFmt);
        cursor.insertBlock();

        QTextCharFormat italicFmt;
        italicFmt.setFontItalic(true);
        italicFmt.setFontPointSize(12);
        cursor.insertText(QStringLiteral("Italic text"), italicFmt);
        cursor.insertBlock();

        QTextCharFormat underlineFmt;
        underlineFmt.setFontUnderline(true);
        underlineFmt.setFontPointSize(12);
        cursor.insertText(QStringLiteral("Underlined text"), underlineFmt);
        cursor.insertBlock();

        QTextListFormat listFmt;
        listFmt.setStyle(QTextListFormat::ListDisc);
        listFmt.setIndent(1);
        cursor.insertText(QStringLiteral("Bullet item 1"));
        cursor.createList(listFmt);
        cursor.insertBlock();
        cursor.insertText(QStringLiteral("Bullet item 2"));
        cursor.insertBlock();
        cursor.insertText(QStringLiteral("Bullet item 3"));

        QMarginsF margins(72, 72, 72, 72);
        QColor bg(Qt::white);

        QByteArray docx = DocxConverter::saveToDocx(&doc, margins, bg);
        QVERIFY(!docx.isEmpty());
        QVERIFY(docx.size() > 100);

        QVERIFY(docx.size() >= 4);
        QCOMPARE(docx[0], 'P');
        QCOMPARE(docx[1], 'K');

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool loadOk = DocxConverter::loadFromDocx(docx, &doc2, loadedMargins, loadedBg);
        QVERIFY(loadOk);
        QString t = doc2.toPlainText();
        QVERIFY(t.contains(QStringLiteral("Bold Heading")));
        QVERIFY(t.contains(QStringLiteral("Italic text")));
        QVERIFY(t.contains(QStringLiteral("Underlined text")));
        QVERIFY(t.contains(QStringLiteral("Bullet item 1")));
        QVERIFY(t.contains(QStringLiteral("Bullet item 2")));
        QVERIFY(t.contains(QStringLiteral("Bullet item 3")));
    }

    void testBoldFormatting()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        QTextCharFormat fmt;
        fmt.setFontWeight(QFont::Bold);
        fmt.setFontPointSize(16);
        cursor.insertText(QStringLiteral("Bold text"), fmt);

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);
        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Bold text")));
    }

    void testFormattingRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QTextCharFormat boldFmt;
        boldFmt.setFontWeight(QFont::Bold);
        boldFmt.setFontPointSize(16);
        cursor.insertText(QStringLiteral("Bold"), boldFmt);
        cursor.insertBlock();

        QTextCharFormat italicFmt;
        italicFmt.setFontItalic(true);
        italicFmt.setFontPointSize(14);
        cursor.insertText(QStringLiteral("Italic"), italicFmt);
        cursor.insertBlock();

        QTextCharFormat unFmt;
        unFmt.setFontUnderline(true);
        cursor.insertText(QStringLiteral("Underline"), unFmt);
        cursor.insertBlock();

        QTextCharFormat strikeFmt;
        strikeFmt.setFontStrikeOut(true);
        cursor.insertText(QStringLiteral("Strikethrough"), strikeFmt);
        cursor.insertBlock();

        QTextCharFormat subFmt;
        subFmt.setVerticalAlignment(QTextCharFormat::AlignSubScript);
        cursor.insertText(QStringLiteral("Subscript"), subFmt);
        cursor.insertBlock();

        QTextCharFormat supFmt;
        supFmt.setVerticalAlignment(QTextCharFormat::AlignSuperScript);
        cursor.insertText(QStringLiteral("Superscript"), supFmt);
        cursor.insertBlock();

        QTextCharFormat sizeFmt;
        sizeFmt.setFontPointSize(24);
        cursor.insertText(QStringLiteral("Large text"), sizeFmt);
        cursor.insertBlock();

        QTextCharFormat colFmt;
        colFmt.setForeground(QColor(Qt::red));
        cursor.insertText(QStringLiteral("Red text"), colFmt);
        cursor.insertBlock();

        QTextBlockFormat centerFmt;
        centerFmt.setAlignment(Qt::AlignCenter);
        cursor.insertBlock(centerFmt);
        cursor.insertText(QStringLiteral("Centered"));

        QMarginsF margins(72, 72, 72, 72);
        QByteArray docx = DocxConverter::saveToDocx(&doc, margins);
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);

        QString text = doc2.toPlainText();
        QVERIFY(text.contains(QStringLiteral("Bold")));
        QVERIFY(text.contains(QStringLiteral("Italic")));
        QVERIFY(text.contains(QStringLiteral("Underline")));
        QVERIFY(text.contains(QStringLiteral("Strikethrough")));
        QVERIFY(text.contains(QStringLiteral("Subscript")));
        QVERIFY(text.contains(QStringLiteral("Superscript")));
        QVERIFY(text.contains(QStringLiteral("Large")));
        QVERIFY(text.contains(QStringLiteral("Red")));
        QVERIFY(text.contains(QStringLiteral("Centered")));
    }

    void testListRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QTextListFormat listFmt;
        listFmt.setStyle(QTextListFormat::ListDisc);
        listFmt.setIndent(1);
        cursor.insertText(QStringLiteral("Item A"));
        cursor.createList(listFmt);
        cursor.insertBlock();
        cursor.insertText(QStringLiteral("Item B"));
        cursor.insertBlock();
        cursor.insertText(QStringLiteral("Item C"));

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);

        QString text = doc2.toPlainText().trimmed();
        QVERIFY(text.contains(QStringLiteral("Item A")));
        QVERIFY(text.contains(QStringLiteral("Item B")));
        QVERIFY(text.contains(QStringLiteral("Item C")));
    }

    void testPageSizeAndMargins()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        QTextCursor cursor(&doc);
        cursor.insertText(QStringLiteral("Test"));

        QMarginsF margins(48, 48, 48, 48);
        QByteArray docx = DocxConverter::saveToDocx(&doc, margins);
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);
        QCOMPARE(doc2.pageSize().width(), 816.0);
        QCOMPARE(doc2.pageSize().height(), 1056.0);
        QVERIFY(loadedMargins.left() > 0);
        QVERIFY(loadedMargins.top() > 0);
    }

    void testEmptyDocument()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(612, 792));
        QMarginsF margins;

        QByteArray docx = DocxConverter::saveToDocx(&doc, margins);
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);
        QCOMPARE(doc2.pageSize().width(), 612.0);
        QCOMPARE(doc2.pageSize().height(), 792.0);
    }

    void testInvalidData()
    {
        QTextDocument doc;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(QByteArrayLiteral("not a zip file"), &doc, margins, bg);
        QVERIFY(!ok);
    }

    void testTableRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QTextTable *table = cursor.insertTable(3, 2);

        for (int r = 0; r < 3; ++r) {
            for (int c = 0; c < 2; ++c) {
                QTextTableCell cell = table->cellAt(r, c);
                QTextCursor cc = cell.firstCursorPosition();
                cc.movePosition(QTextCursor::NextBlock, QTextCursor::KeepAnchor);
                cc.removeSelectedText();
                cc.insertText(QStringLiteral("Cell %1,%2").arg(r).arg(c));
            }
        }

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);

        QString text = doc2.toPlainText();
        QVERIFY(text.contains(QStringLiteral("Cell 0,0")));
        QVERIFY(text.contains(QStringLiteral("Cell 0,1")));
        QVERIFY(text.contains(QStringLiteral("Cell 1,0")));
        QVERIFY(text.contains(QStringLiteral("Cell 2,1")));
    }

    void testMultiLevelListRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QTextListFormat lf1;
        lf1.setStyle(QTextListFormat::ListDecimal);
        lf1.setIndent(1);
        cursor.insertText(QStringLiteral("Level 0 - Item 1"));
        cursor.createList(lf1);
        cursor.insertBlock();

        QTextListFormat lf2;
        lf2.setStyle(QTextListFormat::ListLowerAlpha);
        lf2.setIndent(2);
        cursor.insertText(QStringLiteral("Level 1 - Item A"));
        cursor.createList(lf2);
        cursor.insertBlock();

        cursor.insertText(QStringLiteral("Level 1 - Item B"));

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);

        QString text = doc2.toPlainText();
        QVERIFY(text.contains(QStringLiteral("Level 0 - Item 1")));
        QVERIFY(text.contains(QStringLiteral("Level 1 - Item A")));
        QVERIFY(text.contains(QStringLiteral("Level 1 - Item B")));
    }

    void testFontFamiliesRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QTextCharFormat fontFmt1;
        fontFmt1.setFontFamilies(QStringList{QStringLiteral("Arial"), QStringLiteral("Helvetica")});
        fontFmt1.setFontPointSize(12);
        cursor.insertText(QStringLiteral("Arial text"), fontFmt1);
        cursor.insertBlock();

        QTextCharFormat fontFmt2;
        fontFmt2.setFontFamilies(QStringList{QStringLiteral("Times New Roman")});
        fontFmt2.setFontPointSize(14);
        cursor.insertText(QStringLiteral("Times text"), fontFmt2);

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);

        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Arial text")));
        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Times text")));
    }

    void testPageBackground()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        cursor.insertText(QStringLiteral("Background test"));

        QMarginsF margins(72, 72, 72, 72);
        QColor bg(QColor(240, 240, 240));

        QByteArray docx = DocxConverter::saveToDocx(&doc, margins, bg);
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF loadedMargins;
        QColor loadedBg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, loadedMargins, loadedBg);
        QVERIFY(ok);
        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Background test")));
    }

    void testImageRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QImage img(20, 20, QImage::Format_ARGB32);
        img.fill(Qt::blue);
        QTextImageFormat imgFmt;
        QString imgName = QStringLiteral("test_image.png");
        doc.addResource(QTextDocument::ImageResource, QUrl(imgName), img);
        imgFmt.setName(imgName);
        imgFmt.setWidth(40);
        imgFmt.setHeight(40);
        cursor.insertImage(imgFmt);
        cursor.insertBlock();
        cursor.insertText(QStringLiteral("After image"));

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);
        QVERIFY(doc2.toPlainText().contains(QStringLiteral("After image")));
    }

    void testTabCharacter()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        cursor.insertText(QStringLiteral("Before\tAfter"));

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);
        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Before\tAfter")));
    }
};

QTEST_MAIN(TestDocxConverter)
#include "test_DocxConverter.moc"
