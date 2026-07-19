#include <QTest>
#include <QSignalSpy>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextCharFormat>
#include <QTextBlock>
#include <QFont>
#include <QTemporaryFile>
#include <QFile>
#include "editor/DocumentEditor.h"

class TestDocumentEditor : public QObject
{
    Q_OBJECT

private slots:
    void testInitialState()
    {
        DocumentEditor editor;
        QVERIFY(editor.documentName().isEmpty());
        QVERIFY(!editor.isModified());
        QVERIFY(!editor.isReadOnlyFile());
        QCOMPARE(editor.zoomLevel(), 1.0);
        QCOMPARE(editor.lineCount(), 1);
        QCOMPARE(editor.wordCount(), 0);
    }

    void testDocumentName()
    {
        DocumentEditor editor;
        editor.setDocumentName(QStringLiteral("/tmp/test.xaml"));
        QCOMPARE(editor.documentName(), QStringLiteral("/tmp/test.xaml"));
    }

    void testFileChangedSignal()
    {
        DocumentEditor editor;
        QSignalSpy spy(&editor, &DocumentEditor::modifiedChanged);
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("hello"));
        QCOMPARE(spy.count(), 1);
        QCOMPARE(spy.takeFirst().at(0).toBool(), true);
        QVERIFY(editor.isModified());
    }

    void testToggleBold()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleBold();
        QVERIFY(editor.textCursor().charFormat().fontWeight() >= QFont::Bold);

        editor.toggleBold();
        QVERIFY(editor.textCursor().charFormat().fontWeight() < QFont::Bold);
    }

    void testToggleItalic()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleItalic();
        QVERIFY(editor.textCursor().charFormat().fontItalic());

        editor.toggleItalic();
        QVERIFY(!editor.textCursor().charFormat().fontItalic());
    }

    void testToggleUnderline()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleUnderline();
        QVERIFY(editor.textCursor().charFormat().fontUnderline());

        editor.toggleUnderline();
        QVERIFY(!editor.textCursor().charFormat().fontUnderline());
    }

    void testToggleStrikethrough()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleStrikethrough();
        QVERIFY(editor.textCursor().charFormat().fontStrikeOut());

        editor.toggleStrikethrough();
        QVERIFY(!editor.textCursor().charFormat().fontStrikeOut());
    }

    void testToggleSubscript()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleSubscript();
        QCOMPARE(editor.textCursor().charFormat().verticalAlignment(),
                 QTextCharFormat::AlignSubScript);

        editor.toggleSubscript();
        QCOMPARE(editor.textCursor().charFormat().verticalAlignment(),
                 QTextCharFormat::AlignNormal);
    }

    void testToggleSuperscript()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleSuperscript();
        QCOMPARE(editor.textCursor().charFormat().verticalAlignment(),
                 QTextCharFormat::AlignSuperScript);

        editor.toggleSuperscript();
        QCOMPARE(editor.textCursor().charFormat().verticalAlignment(),
                 QTextCharFormat::AlignNormal);
    }

    void testSubscriptClearsSuperscript()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleSuperscript();
        editor.toggleSubscript();
        QCOMPARE(editor.textCursor().charFormat().verticalAlignment(),
                 QTextCharFormat::AlignSubScript);
    }

    void testSuperscriptClearsSubscript()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::Document);
        editor.setTextCursor(cursor);

        editor.toggleSubscript();
        editor.toggleSuperscript();
        QCOMPARE(editor.textCursor().charFormat().verticalAlignment(),
                 QTextCharFormat::AlignSuperScript);
    }

    void testClearFormatting()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world"));
        cursor.select(QTextCursor::WordUnderCursor);
        editor.setTextCursor(cursor);

        editor.toggleBold();
        editor.toggleItalic();
        editor.clearFormatting();

        QTextCharFormat fmt = editor.textCursor().charFormat();
        QVERIFY(fmt.fontWeight() < QFont::Bold);
        QVERIFY(!fmt.fontItalic());
    }

    void testZoomLevel()
    {
        DocumentEditor editor;
        editor.setBaseFontPointSize(12);
        editor.setZoomLevel(2.0);
        QCOMPARE(editor.zoomLevel(), 2.0);

        // Zoom is clamped
        editor.setZoomLevel(10.0);
        QCOMPARE(editor.zoomLevel(), 5.0);

        editor.setZoomLevel(0.05);
        QCOMPARE(editor.zoomLevel(), 0.1);
    }

    void testWordCount()
    {
        DocumentEditor editor;
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("hello world foo bar"));
        editor.refreshStats();
        QCOMPARE(editor.wordCount(), 4);
    }

    void testLineCount()
    {
        DocumentEditor editor;
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("line1\nline2\nline3"));
        editor.refreshStats();
        QCOMPARE(editor.lineCount(), 3);
    }

    void testGoToLine()
    {
        DocumentEditor editor;
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("line1\nline2\nline3"));

        editor.goToLine(2);
        editor.refreshStats();
        QCOMPARE(editor.selectedLineNumber(), 2);
    }

    void testAlignment()
    {
        DocumentEditor editor;
        editor.setParagraphAlignment(Qt::AlignCenter);
        QCOMPARE(editor.textCursor().blockFormat().alignment(), Qt::AlignCenter);

        editor.setParagraphAlignment(Qt::AlignRight);
        QCOMPARE(editor.textCursor().blockFormat().alignment(), Qt::AlignRight);
    }

    void testIndentMoreLess()
    {
        DocumentEditor editor;
        int initial = editor.textCursor().blockFormat().indent();
        editor.indentMore();
        QCOMPARE(editor.textCursor().blockFormat().indent(), initial + 1);
        editor.indentLess();
        QCOMPARE(editor.textCursor().blockFormat().indent(), initial);
    }

    void testToggleBulletList()
    {
        DocumentEditor editor;
        editor.toggleBulletList();
        QVERIFY(editor.textCursor().block().textList() != nullptr);
        editor.toggleBulletList();
        QVERIFY(editor.textCursor().block().textList() == nullptr);
    }

    void testToggleNumberList()
    {
        DocumentEditor editor;
        editor.toggleNumberList();
        QVERIFY(editor.textCursor().block().textList() != nullptr);
        editor.toggleNumberList();
        QVERIFY(editor.textCursor().block().textList() == nullptr);
    }

    void testPageSize()
    {
        DocumentEditor editor;
        editor.setPageWidth(800);
        editor.setPageHeight(600);
        QCOMPARE(editor.pageWidth(), 800.0);
        QCOMPARE(editor.pageHeight(), 600.0);
    }

    void testPageMargins()
    {
        DocumentEditor editor;
        QMarginsF margins(50, 50, 50, 50);
        editor.setPageMargins(margins);
        QCOMPARE(editor.pageMargins().left(), 50.0);
        QCOMPARE(editor.pageMargins().top(), 50.0);
    }

    void testUpperCase()
    {
        DocumentEditor editor;
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("hello"));
        cursor.select(QTextCursor::WordUnderCursor);
        editor.toUpperCase();
        QCOMPARE(editor.toPlainText().trimmed(), QStringLiteral("HELLO"));
    }

    void testLowerCase()
    {
        DocumentEditor editor;
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("HELLO"));
        cursor.select(QTextCursor::WordUnderCursor);
        editor.toLowerCase();
        QCOMPARE(editor.toPlainText().trimmed(), QStringLiteral("hello"));
    }

    void testTextDirection()
    {
        DocumentEditor editor;
        editor.setTextDirection(Qt::RightToLeft);
        QCOMPARE(editor.textCursor().blockFormat().layoutDirection(), Qt::RightToLeft);
    }

    void testPageOrientationToggle()
    {
        DocumentEditor editor;
        qreal w = editor.pageWidth();
        qreal h = editor.pageHeight();
        editor.togglePageOrientation();
        QCOMPARE(editor.pageWidth(), h);
        QCOMPARE(editor.pageHeight(), w);
    }

    void testFindWord()
    {
        DocumentEditor editor;
        QTextCursor cursor = editor.textCursor();
        cursor.insertText(QStringLiteral("hello world hello"));
        cursor.setPosition(0);
        editor.setTextCursor(cursor);

        QTextCursor found = editor.findWord(QStringLiteral("world"));
        QVERIFY(!found.isNull());
        QCOMPARE(found.selectedText(), QStringLiteral("world"));
    }

    void testSaveAndLoadXaml()
    {
        DocumentEditor editor;
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("Hello XAML"));

        QTemporaryFile tmpFile;
        tmpFile.open();
        QString path = tmpFile.fileName() + QStringLiteral(".xaml");
        tmpFile.close();

        QVERIFY(editor.saveToFile(path));
        QVERIFY(editor.isModified() == false);

        DocumentEditor editor2;
        QVERIFY(editor2.loadFromFile(path));
        QCOMPARE(editor2.toPlainText().trimmed(), QStringLiteral("Hello XAML"));

        QFile::remove(path);
    }

    void testSaveAndLoadTxt()
    {
        DocumentEditor editor;
        QTextCursor cursor(editor.document());
        cursor.insertText(QStringLiteral("Hello Text"));

        QTemporaryFile tmpFile;
        tmpFile.open();
        QString path = tmpFile.fileName() + QStringLiteral(".txt");
        tmpFile.close();

        QVERIFY(editor.saveToFile(path));

        DocumentEditor editor2;
        QVERIFY(editor2.loadFromFile(path));
        QCOMPARE(editor2.toPlainText().trimmed(), QStringLiteral("Hello Text"));

        QFile::remove(path);
    }

    void testPageBackground()
    {
        DocumentEditor editor;
        QColor bg(Qt::yellow);
        editor.setPageBackground(bg);
        QCOMPARE(editor.pageBackground(), bg);
    }
};

QTEST_MAIN(TestDocumentEditor)
#include "test_DocumentEditor.moc"
