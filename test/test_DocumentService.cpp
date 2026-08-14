#include <QTest>
#include <QSignalSpy>
#include <QMainWindow>
#include <QMdiArea>
#include <QStatusBar>
#include <QWidget>
#include <QTemporaryFile>
#include <QTemporaryDir>
#include <QDir>
#include <QFile>
#include <QImage>
#include <QSet>
#include <QTextBlock>
#include <QTextFragment>
#include "services/DocumentService.h"
#include "editor/DocumentTab.h"
#include "editor/DocumentManager.h"
#include "editor/DocumentEditor.h"

class TestDocumentService : public QObject
{
    Q_OBJECT

private:
    QMainWindow *m_mainWindow = nullptr;
    QMdiArea *m_mdiArea = nullptr;
    DocumentManager *m_docManager = nullptr;
    DocumentService *m_service = nullptr;

private slots:
    void initTestCase()
    {
        m_mainWindow = new QMainWindow();
        m_mdiArea = new QMdiArea(m_mainWindow);
        m_mainWindow->setCentralWidget(m_mdiArea);
        m_docManager = new DocumentManager(m_mdiArea, m_mainWindow, m_mainWindow);

        m_service = new DocumentService(m_docManager, m_mainWindow);
    }

    void cleanupTestCase()
    {
        delete m_mainWindow;
    }

    void cleanup()
    {
        for (DocumentTab *tab : m_docManager->allTabs())
            tab->editor()->setModified(false);
        m_docManager->closeAllDocuments();
    }

    void testNewDocument()
    {
        QCOMPARE(m_docManager->tabCount(), 0);
        m_service->newDocument();
        QCOMPARE(m_docManager->tabCount(), 1);
        QVERIFY(m_docManager->activeEditor() != nullptr);
    }

    void testNewDocumentSignal()
    {
        QSignalSpy spy(m_service, &DocumentService::documentOpened);
        m_service->newDocument();
        QCOMPARE(spy.count(), 1);
        QCOMPARE(spy.takeFirst().at(0).toString(), QString());
    }

    void testEmbedImageInDocumentInvalidPath()
    {
        m_service->newDocument();
        DocumentEditor *editor = m_docManager->activeEditor();
        QVERIFY(editor != nullptr);

        // Embedding a non-existent image should be a no-op
        DocumentService::embedImageInDocument(editor,
            QStringLiteral("/nonexistent/image.png"));
        QVERIFY(editor->toPlainText().isEmpty());
    }

    void testImportImageNoEditor()
    {
        // Should not crash and should not create an editor
        m_service->importImage();
        QCOMPARE(m_docManager->tabCount(), 0);
    }

    void testExportImageWithNoEditor()
    {
        QCOMPARE(m_docManager->tabCount(), 0);
        m_service->exportImage();
        QCOMPARE(m_docManager->tabCount(), 0);
    }

    void testExportPdfWithNoEditor()
    {
        m_service->exportPdf();
        QCOMPARE(m_docManager->tabCount(), 0);
    }

    void testPrintDocumentWithNoEditor()
    {
        m_service->printDocument();
        QCOMPARE(m_docManager->tabCount(), 0);
    }

    void testPageSetupWithNoEditor()
    {
        m_service->pageSetup();
        QCOMPARE(m_docManager->tabCount(), 0);
    }

    void testRevertDocumentWithNoEditor()
    {
        m_service->revertDocument();
        QCOMPARE(m_docManager->tabCount(), 0);
    }

    void testRevertDocumentWithUnsavedEditor()
    {
        m_service->newDocument();
        DocumentEditor *editor = m_docManager->activeEditor();
        QVERIFY(editor != nullptr);

        // Reverting an unsaved new document should be a no-op
        m_service->revertDocument();
        QCOMPARE(m_docManager->tabCount(), 1);
        QVERIFY(m_docManager->activeEditor() == editor);
    }

    void testEmbedImageInDocumentValidPath()
    {
        m_service->newDocument();
        DocumentEditor *editor = m_docManager->activeEditor();
        QVERIFY(editor != nullptr);

        // Create a small valid PNG file
        QTemporaryFile tmpPng;
        QVERIFY(tmpPng.open());
        QString pngPath = tmpPng.fileName() + QStringLiteral(".png");
        tmpPng.close();

        QImage img(10, 10, QImage::Format_ARGB32);
        img.fill(Qt::red);
        QVERIFY(img.save(pngPath));

        DocumentService::embedImageInDocument(editor, pngPath, 50, 50);
        QString html = editor->toHtml();
        QVERIFY(html.contains(QStringLiteral("img")) || html.contains(QStringLiteral("embed")));

        QFile::remove(pngPath);
    }

    void testEmbedImageSameBasenameCollision()
    {
        m_service->newDocument();
        DocumentEditor *editor = m_docManager->activeEditor();
        QVERIFY(editor != nullptr);

        // Two images in different directories that share a basename must not
        // overwrite each other's document resource.
        QTemporaryDir dirA;
        QTemporaryDir dirB;
        QVERIFY(dirA.isValid());
        QVERIFY(dirB.isValid());

        const QString pngA = dirA.path() + QStringLiteral("/photo.png");
        const QString pngB = dirB.path() + QStringLiteral("/photo.png");

        QImage red(10, 10, QImage::Format_ARGB32);
        red.fill(Qt::red);
        QImage green(10, 10, QImage::Format_ARGB32);
        green.fill(Qt::green);
        QVERIFY(red.save(pngA));
        QVERIFY(green.save(pngB));

        DocumentService::embedImageInDocument(editor, pngA, 50, 50);
        DocumentService::embedImageInDocument(editor, pngB, 50, 50);

        // Each fragment must reference a distinct resource name.
        QSet<QString> names;
        for (QTextBlock block = editor->document()->begin();
             block != editor->document()->end(); block = block.next()) {
            for (auto it = block.begin(); !it.atEnd(); ++it) {
                const QTextFragment fragment = it.fragment();
                if (!fragment.isValid())
                    continue;
                const QTextCharFormat fmt = fragment.charFormat();
                if (fmt.isImageFormat())
                    names.insert(fmt.toImageFormat().name());
            }
        }
        QCOMPARE(names.size(), 2);
    }

    void testArchiveToolDetection()
    {
#ifdef Q_OS_UNIX
        // Regression test: real 7z exits 7 and unzip exits 10 for "--version",
        // so probing must use their documented info flags ("7z i", "unzip -v").
        QTemporaryDir dir;
        QVERIFY(dir.isValid());
        const QString binDir = dir.path() + QStringLiteral("/bin");
        QVERIFY(QDir().mkpath(binDir));

        auto writeScript = [&binDir](const QString &name, const QString &body) {
            QFile script(binDir + QLatin1Char('/') + name);
            QVERIFY(script.open(QIODevice::WriteOnly));
            script.write(body.toUtf8());
            script.close();
            QVERIFY(script.setPermissions(
                QFileDevice::ReadOwner | QFileDevice::WriteOwner | QFileDevice::ExeOwner));
        };

        // Mimic the real tools: exit non-zero for --version, zero for the
        // info flags used by the probe.
        writeScript(QStringLiteral("7z"), QStringLiteral("#!/bin/sh\n"
            "[ \"$1\" = \"i\" ] && exit 0\nexit 7\n"));
        writeScript(QStringLiteral("unzip"), QStringLiteral("#!/bin/sh\n"
            "[ \"$1\" = \"-v\" ] && exit 0\nexit 10\n"));
        writeScript(QStringLiteral("zip"), QStringLiteral("#!/bin/sh\n"
            "[ \"$1\" = \"--version\" ] && exit 0\nexit 1\n"));

        const QByteArray savedPath = qgetenv("PATH");
        const QByteArray testPath = QFile::encodeName(binDir);
        qputenv("PATH", testPath);

        QCOMPARE(DocumentService::detectArchiver(), QStringLiteral("7z"));
        QCOMPARE(DocumentService::detectCompressor(), QStringLiteral("7z"));

        // Remove 7z; unzip must be found as the archiver and zip as compressor.
        QVERIFY(QFile::remove(binDir + QStringLiteral("/7z")));
        QCOMPARE(DocumentService::detectArchiver(), QStringLiteral("unzip"));
        QCOMPARE(DocumentService::detectCompressor(), QStringLiteral("zip"));

        // Remove the remaining tools; nothing may be detected.
        QVERIFY(QFile::remove(binDir + QStringLiteral("/unzip")));
        QVERIFY(QFile::remove(binDir + QStringLiteral("/zip")));
        QCOMPARE(DocumentService::detectArchiver(), QString());
        QCOMPARE(DocumentService::detectCompressor(), QString());

        qputenv("PATH", savedPath);
#endif
    }

};

QTEST_MAIN(TestDocumentService)
#include "test_DocumentService.moc"
