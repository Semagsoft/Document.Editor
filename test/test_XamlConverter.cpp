#include <QTest>
#include <QTextDocument>
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
};

QTEST_MAIN(TestXamlConverter)
#include "test_XamlConverter.moc"
