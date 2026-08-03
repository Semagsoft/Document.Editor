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
#include <QFile>

#include "converters/DocxConverter.h"

class TestDocxConverter : public QObject
{
    Q_OBJECT

private:
    static QTextTable *findFirstTable(QTextDocument *doc)
    {
        for (QTextBlock block = doc->begin(); block.isValid(); block = block.next()) {
            QTextFrame *frame = doc->frameAt(block.position());
            while (frame) {
                if (QTextTable *table = qobject_cast<QTextTable *>(frame))
                    return table;
                frame = frame->parentFrame();
            }
        }
        return nullptr;
    }

    static int countImageFragments(QTextDocument *doc)
    {
        int count = 0;
        for (QTextBlock block = doc->begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                if (it.fragment().charFormat().isImageFormat())
                    ++count;
            }
        }
        return count;
    }

    static QString findAnchorHref(QTextDocument *doc)
    {
        for (QTextBlock block = doc->begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                QTextCharFormat cf = it.fragment().charFormat();
                if (cf.isAnchor() && !cf.anchorHref().isEmpty())
                    return cf.anchorHref();
            }
        }
        return QString();
    }

    static bool hasListStyle(QTextDocument *doc, QTextListFormat::Style style)
    {
        for (QTextBlock block = doc->begin(); block.isValid(); block = block.next()) {
            QTextList *list = block.textList();
            if (list && list->format().style() == style)
                return true;
        }
        return false;
    }

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
        QVERIFY(hasListStyle(&doc2, QTextListFormat::ListDisc));
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

        QTextTable *loaded = findFirstTable(&doc2);
        QVERIFY(loaded != nullptr);
        QCOMPARE(loaded->rows(), 3);
        QCOMPARE(loaded->columns(), 2);
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
        QVERIFY(hasListStyle(&doc2, QTextListFormat::ListDecimal));
        QVERIFY(hasListStyle(&doc2, QTextListFormat::ListLowerAlpha));
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
        QCOMPARE(countImageFragments(&doc2), 1);
        for (QTextBlock block = doc2.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                if (it.fragment().charFormat().isImageFormat()) {
                    QTextImageFormat ifmt = it.fragment().charFormat().toImageFormat();
                    QCOMPARE(qRound(ifmt.width()), 40);
                    QCOMPARE(qRound(ifmt.height()), 40);
                }
            }
        }
    }

    void testHyperlinkRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QTextCharFormat linkFmt;
        linkFmt.setAnchor(true);
        linkFmt.setAnchorHref(QStringLiteral("https://example.com"));
        linkFmt.setForeground(QColor(Qt::blue));
        linkFmt.setFontUnderline(true);
        cursor.insertText(QStringLiteral("Click here"), linkFmt);

        cursor.insertBlock();
        QTextCharFormat normalFmt;
        cursor.insertText(QStringLiteral("Normal text"), normalFmt);

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);

        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Click here")));
        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Normal text")));
        QCOMPARE(findAnchorHref(&doc2), QStringLiteral("https://example.com"));
    }

    void testMergedCellsRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);
        QTextTable *table = cursor.insertTable(2, 3);

        table->mergeCells(0, 0, 1, 2);

        for (int r = 0; r < 2; ++r) {
            for (int c = 0; c < 3; ++c) {
                QTextTableCell cell = table->cellAt(r, c);
                if (cell.row() != r || cell.column() != c) continue;
                QTextCursor cc = cell.firstCursorPosition();
                cc.movePosition(QTextCursor::NextBlock, QTextCursor::KeepAnchor);
                cc.removeSelectedText();
                cc.insertText(QStringLiteral("R%1C%2").arg(r).arg(c));
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
        QVERIFY(text.contains(QStringLiteral("R0C0")));
        QVERIFY(text.contains(QStringLiteral("R1C1")));

        QTextTable *loaded = findFirstTable(&doc2);
        QVERIFY(loaded != nullptr);
        QCOMPARE(loaded->rows(), 2);
        QCOMPARE(loaded->columns(), 3);
        QCOMPARE(loaded->cellAt(0, 0).columnSpan(), 2);
    }

    void testVerticalMergeRoundTrip()
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
                cc.insertText(QStringLiteral("R%1C%2").arg(r).arg(c));
            }
        }

        table->mergeCells(0, 0, 2, 1);
        QCOMPARE(table->cellAt(0, 0).rowSpan(), 2);

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);

        QTextTable *loadedTable = findFirstTable(&doc2);
        QVERIFY(loadedTable != nullptr);
        QCOMPARE(loadedTable->rows(), 3);
        QCOMPARE(loadedTable->columns(), 2);
        QCOMPARE(loadedTable->cellAt(0, 0).rowSpan(), 2);
    }

    void testMultipleImagesSameParagraph()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QImage img1(20, 20, QImage::Format_ARGB32);
        img1.fill(Qt::blue);
        QTextImageFormat imgFmt1;
        QString name1 = QStringLiteral("img_one.png");
        doc.addResource(QTextDocument::ImageResource, QUrl(name1), img1);
        imgFmt1.setName(name1);
        imgFmt1.setWidth(40);
        imgFmt1.setHeight(40);

        QImage img2(20, 20, QImage::Format_ARGB32);
        img2.fill(Qt::red);
        QTextImageFormat imgFmt2;
        QString name2 = QStringLiteral("img_two.png");
        doc.addResource(QTextDocument::ImageResource, QUrl(name2), img2);
        imgFmt2.setName(name2);
        imgFmt2.setWidth(40);
        imgFmt2.setHeight(40);

        cursor.insertImage(imgFmt1);
        cursor.insertImage(imgFmt2);
        cursor.insertBlock();
        cursor.insertText(QStringLiteral("After images"));

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);
        QVERIFY(doc2.toPlainText().contains(QStringLiteral("After images")));

        int imageCount = 0;
        for (QTextBlock block = doc2.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                if (it.fragment().charFormat().isImageFormat())
                    ++imageCount;
            }
        }
        QCOMPARE(imageCount, 2);
    }

    void testNamedStyleRoundTrip()
    {
        QTextDocument doc;
        QTextCursor cursor(&doc);

        QTextBlockFormat headingFmt;
        headingFmt.setProperty(QTextFormat::UserProperty, QStringLiteral("Heading1"));
        headingFmt.setAlignment(Qt::AlignCenter);
        cursor.insertBlock(headingFmt);
        cursor.insertText(QStringLiteral("Heading Text"));

        cursor.insertBlock();
        cursor.insertText(QStringLiteral("Normal Text"));

        QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());
        QVERIFY(!docx.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(docx, &doc2, margins, bg);
        QVERIFY(ok);

        QVERIFY(doc2.toPlainText().contains(QStringLiteral("Heading Text")));

        bool foundHeading = false;
        for (QTextBlock block = doc2.begin(); block.isValid(); block = block.next()) {
            if (block.text() == QStringLiteral("Heading Text")) {
                foundHeading = true;
                QCOMPARE(block.blockFormat().property(QTextFormat::UserProperty).toString(),
                         QStringLiteral("Heading1"));
            }
        }
        QVERIFY(foundHeading);
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

    void testRealFixtureLoads()
    {
        QFile file(QStringLiteral(DOCX_FIXTURE_DIR) + QStringLiteral("/merge_and_textbox.docx"));
        QVERIFY(file.open(QIODevice::ReadOnly));
        QByteArray data = file.readAll();
        QVERIFY(!data.isEmpty());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        bool ok = DocxConverter::loadFromDocx(data, &doc2, margins, bg);
        QVERIFY(ok);

        QString text = doc2.toPlainText();
        QVERIFY(text.contains(QStringLiteral("Fixture title")));
        QVERIFY(text.contains(QStringLiteral("A1")));
        QVERIFY(text.contains(QStringLiteral("B2")));
        QVERIFY(text.contains(QStringLiteral("C3")));
        QVERIFY(text.contains(QStringLiteral("D4")));

        QTextTable *table = findFirstTable(&doc2);
        QVERIFY(table != nullptr);
        QCOMPARE(table->rows(), 3);
        QCOMPARE(table->columns(), 2);
        QCOMPARE(table->cellAt(0, 0).rowSpan(), 2);
        QCOMPARE(table->cellAt(0, 1).rowSpan(), 1);

        QByteArray resaved = DocxConverter::saveToDocx(&doc2, margins);
        QVERIFY(!resaved.isEmpty());
        QTextDocument doc3;
        QMarginsF m3;
        QColor bg3;
        QVERIFY(DocxConverter::loadFromDocx(resaved, &doc3, m3, bg3));
        QTextTable *table3 = findFirstTable(&doc3);
        QVERIFY(table3 != nullptr);
        QCOMPARE(table3->cellAt(0, 0).rowSpan(), 2);
        QCOMPARE(table3->cellAt(0, 1).rowSpan(), 1);
    }
};

QTEST_MAIN(TestDocxConverter)
#include "test_DocxConverter.moc"
