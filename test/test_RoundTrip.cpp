// Regression tests for the bugs found in the code review:
//  * DOCX font size / subscript-superscript / paragraph alignment were lost
//    on save -> reload (writer/reader OOXML attribute-vs-text mismatch).
//  * XAML LineHeight was stored as a percentage but re-read as a fixed
//    pixel distance (100x distortion).
//  * XAML character formatting inside list items was dropped on load.
//  * Embedded images vanished from the live document after open (fragment
//    copy does not carry QTextDocument resources).
//  * Page margins were zeroed on open (temp document never got the root
//    frame format).
//  * isSafeFilePath had no test coverage.
#include <QTest>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextImageFormat>
#include <QTemporaryFile>
#include <QFile>
#include <QMarginsF>
#include <QColor>

#include "converters/XamlConverter.h"
#include "converters/DocxConverter.h"
#include "editor/DocumentEditor.h"

class TestRoundTrip : public QObject
{
    Q_OBJECT

private:
    static QTextBlock findBlock(const QTextDocument &doc, const QString &text)
    {
        for (QTextBlock block = doc.begin(); block.isValid(); block = block.next()) {
            if (block.text().contains(text))
                return block;
        }
        return QTextBlock();
    }

private slots:
    void docxPreservesFontSize()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        QTextCursor cursor(&doc);
        QTextCharFormat fmt;
        fmt.setFontPointSize(22.0);
        cursor.insertText(QStringLiteral("sized text"), fmt);

        const QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        QVERIFY(DocxConverter::loadFromDocx(docx, &doc2, margins, bg));

        const QTextBlock block = findBlock(doc2, QStringLiteral("sized"));
        QVERIFY(block.isValid());
        bool foundSize = false;
        for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
            const QTextFragment f = it.fragment();
            if (f.text().contains(QStringLiteral("sized")))
                foundSize = qFuzzyCompare(f.charFormat().fontPointSize(), 22.0);
        }
        QVERIFY2(foundSize, "font point size must survive a DOCX roundtrip");
    }

    void docxPreservesSubAndSuperscript()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        QTextCursor cursor(&doc);
        QTextCharFormat sub;
        sub.setVerticalAlignment(QTextCharFormat::AlignSubScript);
        QTextCharFormat sup;
        sup.setVerticalAlignment(QTextCharFormat::AlignSuperScript);
        cursor.insertText(QStringLiteral("H"), QTextCharFormat());
        cursor.insertText(QStringLiteral("2"), sub);
        cursor.insertText(QStringLiteral("O X"), QTextCharFormat());
        cursor.insertText(QStringLiteral("2"), sup);

        const QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        QVERIFY(DocxConverter::loadFromDocx(docx, &doc2, margins, bg));

        bool sawSub = false, sawSuper = false;
        for (QTextBlock block = doc2.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                const QTextFragment f = it.fragment();
                if (f.charFormat().verticalAlignment() == QTextCharFormat::AlignSubScript)
                    sawSub = true;
                if (f.charFormat().verticalAlignment() == QTextCharFormat::AlignSuperScript)
                    sawSuper = true;
            }
        }
        QVERIFY2(sawSub, "subscript must survive a DOCX roundtrip");
        QVERIFY2(sawSuper, "superscript must survive a DOCX roundtrip");
    }

    void docxPreservesParagraphAlignment()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        QTextCursor cursor(&doc);
        QTextBlockFormat centered;
        centered.setAlignment(Qt::AlignCenter);
        cursor.insertBlock(centered);
        cursor.insertText(QStringLiteral("centered para"));

        const QByteArray docx = DocxConverter::saveToDocx(&doc, QMarginsF());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        QVERIFY(DocxConverter::loadFromDocx(docx, &doc2, margins, bg));

        const QTextBlock block = findBlock(doc2, QStringLiteral("centered"));
        QVERIFY(block.isValid());
        QCOMPARE(block.blockFormat().alignment(), Qt::AlignCenter);
    }

    void xamlPreservesProportionalLineHeight()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        QTextCursor cursor(&doc);
        QTextBlockFormat blockFmt;
        blockFmt.setLineHeight(150, QTextBlockFormat::ProportionalHeight);
        cursor.setBlockFormat(blockFmt);
        cursor.insertText(QStringLiteral("spaced paragraph"));

        const QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        QVERIFY(XamlConverter::loadFromXaml(xaml, &doc2, margins, bg));

        const QTextBlock block = findBlock(doc2, QStringLiteral("spaced"));
        QVERIFY(block.isValid());
        QCOMPARE(block.blockFormat().lineHeightType(), QTextBlockFormat::ProportionalHeight);
        QCOMPARE(qRound(block.blockFormat().lineHeight()), 150);
    }

    void xamlPreservesFormattingInListItems()
    {
        QTextDocument doc;
        doc.setPageSize(QSizeF(816, 1056));
        QTextCursor cursor(&doc);
        QTextBlockFormat blockFmt;
        cursor.insertBlock(blockFmt);

        QTextListFormat listFmt;
        listFmt.setStyle(QTextListFormat::ListDisc);
        QTextList *list = cursor.createList(listFmt);
        QVERIFY(list);
        QTextCursor listCursor(list->item(0));
        QTextCharFormat bold;
        bold.setFontWeight(QFont::Bold);
        bold.setFontPointSize(18.0);
        listCursor.insertText(QStringLiteral("bold item"), bold);

        const QString xaml = XamlConverter::saveToXaml(&doc, QMarginsF());

        QTextDocument doc2;
        QMarginsF margins;
        QColor bg;
        QVERIFY(XamlConverter::loadFromXaml(xaml, &doc2, margins, bg));

        bool sawBold = false;
        for (QTextBlock block = doc2.begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                const QTextFragment f = it.fragment();
                if (f.text().contains(QStringLiteral("bold item"))
                    && f.charFormat().fontWeight() == QFont::Bold
                    && qFuzzyCompare(f.charFormat().fontPointSize(), 18.0))
                    sawBold = true;
            }
        }
        QVERIFY2(sawBold, "bold run inside a list item must survive a XAML roundtrip");
    }

    void imagesSurviveSaveLoadSaveCycle()
    {
        QImage image(24, 24, QImage::Format_ARGB32);
        image.fill(Qt::red);

        DocumentEditor editor;
        editor.document()->addResource(QTextDocument::ImageResource,
                                       QUrl(QStringLiteral("test_image")), image);
        QTextImageFormat imgFmt;
        imgFmt.setName(QStringLiteral("test_image"));
        imgFmt.setWidth(24);
        imgFmt.setHeight(24);
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("before image\n"));
        cursor.insertImage(imgFmt);
        cursor.insertText(QStringLiteral("\nafter image"));

        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        const QString path1 = tmp.fileName() + QStringLiteral(".xaml");
        tmp.close();
        QVERIFY(editor.saveToFile(path1));

        DocumentEditor reloaded;
        QVERIFY(reloaded.loadFromFile(path1));

        // The reloaded document must still carry the image: find the image
        // fragment and check its resource resolves (the resource name is
        // rewritten by the loader, so query via the fragment).
        bool sawImage = false;
        bool resourceValid = false;
        for (QTextBlock block = reloaded.document()->begin(); block.isValid(); block = block.next()) {
            for (QTextBlock::iterator it = block.begin(); !it.atEnd(); ++it) {
                const QTextFragment f = it.fragment();
                if (!f.isValid() || !f.charFormat().isImageFormat())
                    continue;
                sawImage = true;
                const QString name = f.charFormat().toImageFormat().name();
                const QVariant res = reloaded.document()->resource(
                    QTextDocument::ImageResource, QUrl(name));
                if (res.isValid())
                    resourceValid = true;
            }
        }
        QVERIFY2(sawImage, "reloaded document must contain the image fragment");
        QVERIFY2(resourceValid, "image resource must survive loadFromFile");

        // And a subsequent save must re-embed it.
        const QString path2 = tmp.fileName() + QStringLiteral("_2.xaml");
        QVERIFY(reloaded.saveToFile(path2, false));
        QFile file(path2);
        QVERIFY(file.open(QIODevice::ReadOnly));
        const QByteArray payload = file.readAll();
        QVERIFY2(payload.contains("base64") || payload.contains("Image"),
                 "re-saved XAML must still embed the image");

        QFile::remove(path1);
        QFile::remove(path2);
    }

    void pageMarginsSurviveLoad()
    {
        DocumentEditor editor;
        editor.setPageMargins(QMarginsF(30, 40, 50, 60));
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("margin test"));

        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        const QString path = tmp.fileName() + QStringLiteral(".xaml");
        tmp.close();
        QVERIFY(editor.saveToFile(path));

        DocumentEditor reloaded;
        // Use deliberately different margins so we can tell who wins.
        reloaded.setPageMargins(QMarginsF(1, 2, 3, 4));
        QVERIFY(reloaded.loadFromFile(path));
        const QTextFrameFormat fmt = reloaded.document()->rootFrame()->frameFormat();
        QCOMPARE(fmt.leftMargin(), 30.0);
        QCOMPARE(fmt.topMargin(), 40.0);
        QCOMPARE(fmt.rightMargin(), 50.0);
        QCOMPARE(fmt.bottomMargin(), 60.0);

        QFile::remove(path);
    }

    void isSafeFilePathBasics()
    {
        QVERIFY(!DocumentEditor::isSafeFilePath(QString()));
        QVERIFY(!DocumentEditor::isSafeFilePath(QStringLiteral("../secret.txt")));
        QVERIFY(!DocumentEditor::isSafeFilePath(QStringLiteral("/tmp/../etc/passwd")));
        QVERIFY(!DocumentEditor::isSafeFilePath(QStringLiteral("a/b/../../c.txt")));
        QVERIFY(DocumentEditor::isSafeFilePath(QStringLiteral("/tmp/normal/file.txt")));
        QVERIFY(DocumentEditor::isSafeFilePath(QStringLiteral("relative/file.txt")));
        // The allowed root constrains absolute paths.
        QVERIFY(DocumentEditor::isSafeFilePath(QStringLiteral("/safe/dir/file.txt"),
                                               QStringLiteral("/safe/dir")));
        QVERIFY(!DocumentEditor::isSafeFilePath(QStringLiteral("/other/dir/file.txt"),
                                                QStringLiteral("/safe/dir")));
        QVERIFY(!DocumentEditor::isSafeFilePath(QStringLiteral("/safe/dir/../other.txt"),
                                                QStringLiteral("/safe/dir")));
    }
};

QTEST_MAIN(TestRoundTrip)
#include "test_RoundTrip.moc"
