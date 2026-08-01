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
};

QTEST_MAIN(TestXamlConverter)
#include "test_XamlConverter.moc"
