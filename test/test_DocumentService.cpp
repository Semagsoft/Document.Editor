#include <QTest>
#include <QSignalSpy>
#include <QMainWindow>
#include <QMdiArea>
#include <QStatusBar>
#include <QWidget>
#include <QTemporaryFile>
#include <QFile>
#include <QImage>
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
        // Should not crash
        m_service->importImage();
        QVERIFY(true);
    }

    void testExportImageWithNoEditor()
    {
        QCOMPARE(m_docManager->tabCount(), 0);
        m_service->exportImage();
        QVERIFY(true);
    }

    void testExportPdfWithNoEditor()
    {
        m_service->exportPdf();
        QVERIFY(true);
    }

    void testPrintDocumentWithNoEditor()
    {
        m_service->printDocument();
        QVERIFY(true);
    }

    void testPageSetupWithNoEditor()
    {
        m_service->pageSetup();
        QVERIFY(true);
    }

    void testRevertDocumentWithNoEditor()
    {
        m_service->revertDocument();
        QVERIFY(true);
    }

    void testRevertDocumentWithUnsavedEditor()
    {
        m_service->newDocument();
        DocumentEditor *editor = m_docManager->activeEditor();
        QVERIFY(editor != nullptr);

        // Reverting an unsaved new document should be a no-op
        m_service->revertDocument();
        QVERIFY(true);
    }

    void testEmbedImageInDocumentValidPath()
    {
        m_service->newDocument();
        DocumentEditor *editor = m_docManager->activeEditor();
        QVERIFY(editor != nullptr);

        // Create a small valid PNG file
        QTemporaryFile tmpPng;
        tmpPng.open();
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

};

QTEST_MAIN(TestDocumentService)
#include "test_DocumentService.moc"
