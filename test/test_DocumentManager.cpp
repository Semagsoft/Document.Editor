#include <QTest>
#include <QMainWindow>
#include <QMdiArea>
#include <QMdiSubWindow>
#include <QTextDocument>
#include <QTemporaryFile>
#include <QFile>
#include <QSignalSpy>
#include "editor/DocumentManager.h"
#include "editor/DocumentTab.h"
#include "editor/DocumentEditor.h"

class TestDocumentManager : public QObject
{
    Q_OBJECT

private:
    QMainWindow *m_mainWindow = nullptr;
    DocumentManager *m_manager = nullptr;
    QMdiArea *m_mdiArea = nullptr;

private slots:
    void initTestCase()
    {
        m_mainWindow = new QMainWindow();
        m_mdiArea = new QMdiArea(m_mainWindow);
        m_mainWindow->setCentralWidget(m_mdiArea);
        m_manager = new DocumentManager(m_mdiArea, m_mainWindow, m_mainWindow);
        qRegisterMetaType<DocumentTab *>("DocumentTab*");
    }

    void cleanupTestCase()
    {
        delete m_mainWindow;
    }

    void cleanup()
    {
        for (DocumentTab *tab : m_manager->allTabs())
            tab->editor()->setModified(false);
        m_manager->closeAllDocuments();
    }

    void testInitialState()
    {
        QCOMPARE(m_manager->tabCount(), 0);
        QVERIFY(m_manager->activeTab() == nullptr);
        QVERIFY(m_manager->activeEditor() == nullptr);
    }

    void testCreateDocument()
    {
        DocumentTab *tab = m_manager->createDocument();
        QVERIFY(tab != nullptr);
        QCOMPARE(m_manager->tabCount(), 1);
        QCOMPARE(m_manager->activeTab(), tab);
        QVERIFY(tab->editor() != nullptr);
    }

    void testCreateMultipleDocuments()
    {
        m_manager->createDocument(QStringLiteral("Doc1"));
        m_manager->createDocument(QStringLiteral("Doc2"));
        QCOMPARE(m_manager->tabCount(), 2);

        QList<DocumentTab *> tabs = m_manager->allTabs();
        QCOMPARE(tabs.size(), 2);
    }

    void testCloseDocument()
    {
        DocumentTab *tab = m_manager->createDocument();
        QCOMPARE(m_manager->tabCount(), 1);

        m_manager->closeDocument(tab);
        QCOMPARE(m_manager->tabCount(), 0);
    }

    void testCloseAllDocuments()
    {
        m_manager->createDocument();
        m_manager->createDocument();
        m_manager->createDocument();
        QCOMPARE(m_manager->tabCount(), 3);

        m_manager->closeAllDocuments();
        QCOMPARE(m_manager->tabCount(), 0);
    }

    void testActiveTab()
    {
        DocumentTab *tab1 = m_manager->createDocument(QStringLiteral("First"));
        DocumentTab *tab2 = m_manager->createDocument(QStringLiteral("Second"));
        Q_UNUSED(tab1);
        Q_UNUSED(tab2);

        QVERIFY(m_manager->activeTab() != nullptr);
        QCOMPARE(m_manager->tabCount(), 2);
    }

    void testActiveEditor()
    {
        m_manager->createDocument();
        DocumentEditor *editor = m_manager->activeEditor();
        QVERIFY(editor != nullptr);
        QVERIFY(qobject_cast<DocumentEditor *>(editor) != nullptr);
    }

    void testEditorInteraction()
    {
        DocumentTab *tab = m_manager->createDocument();
        DocumentEditor *editor = tab->editor();
        editor->textCursor().insertText(QStringLiteral("Hello World"));
        QCOMPARE(editor->toPlainText().trimmed(), QStringLiteral("Hello World"));
    }

    void testAllTabs()
    {
        m_manager->createDocument();
        m_manager->createDocument();
        m_manager->createDocument();

        QList<DocumentTab *> tabs = m_manager->allTabs();
        QCOMPARE(tabs.size(), 3);
    }

    void testSaveFileDirectly()
    {
        DocumentTab *tab = m_manager->createDocument();
        tab->editor()->textCursor().insertText(QStringLiteral("test"));

        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        QString path = tmp.fileName() + QStringLiteral(".txt");
        tmp.close();

        QVERIFY(tab->editor()->saveToFile(path));
        QFile::remove(path);
    }

    void testSaveDocumentWithName()
    {
        DocumentTab *tab = m_manager->createDocument();
        tab->editor()->textCursor().insertText(QStringLiteral("content to save"));

        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        QString path = tmp.fileName() + QStringLiteral(".xaml");
        tmp.close();

        // Give the tab a document name so saveDocument writes directly
        // (no save-as dialog needed).
        tab->setDocumentName(path);
        QVERIFY(m_manager->saveDocument(tab));
        QFile::remove(path);
    }

    void testSaveAllDocuments()
    {
        // Create tabs with document names so saveDocument writes directly
        // (no save-as dialog needed in headless mode).
        DocumentTab *tab1 = m_manager->createDocument(QStringLiteral("Doc 1"));
        Q_UNUSED(tab1);
        tab1->editor()->textCursor().insertText(QStringLiteral("content 1"));

        QTemporaryFile tmp1;
        QVERIFY(tmp1.open());
        QString path1 = tmp1.fileName() + QStringLiteral(".xaml");
        tmp1.close();
        tab1->setDocumentName(path1);

        DocumentTab *tab2 = m_manager->createDocument(QStringLiteral("Doc 2"));
        Q_UNUSED(tab2);
        tab2->editor()->textCursor().insertText(QStringLiteral("content 2"));

        QTemporaryFile tmp2;
        QVERIFY(tmp2.open());
        QString path2 = tmp2.fileName() + QStringLiteral(".xaml");
        tmp2.close();
        tab2->setDocumentName(path2);

        QCOMPARE(m_manager->tabCount(), 2);
        QVERIFY(m_manager->allTabs().first()->editor()->isModified());
        QVERIFY(m_manager->allTabs().at(1)->editor()->isModified());

        m_manager->saveAllDocuments();
        QVERIFY(!m_manager->allTabs().first()->editor()->isModified());
        QVERIFY(!m_manager->allTabs().at(1)->editor()->isModified());

        QFile::remove(path1);
        QFile::remove(path2);
    }

    void testCloseAllButCurrent()
    {
        DocumentTab *tab1 = m_manager->createDocument(QStringLiteral("First"));
        DocumentTab *tab2 = m_manager->createDocument(QStringLiteral("Second"));
        Q_UNUSED(tab1);
        Q_UNUSED(tab2);

        QCOMPARE(m_manager->tabCount(), 2);
        QCOMPARE(m_manager->activeTab(), tab2);

        m_manager->closeAllButCurrent();

        QCOMPARE(m_manager->tabCount(), 1);
        QCOMPARE(m_manager->activeTab(), tab2);
    }

    void testCloseAllButCurrentSingleTab()
    {
        DocumentTab *tab = m_manager->createDocument();
        Q_UNUSED(tab);

        QCOMPARE(m_manager->tabCount(), 1);
        m_manager->closeAllButCurrent();
        QCOMPARE(m_manager->tabCount(), 1);
    }

    void testTabAboutToCloseSignal()
    {
        QSignalSpy spy(m_manager, &DocumentManager::tabAboutToClose);

        DocumentTab *tab = m_manager->createDocument();
        QVERIFY(tab != nullptr);

        m_manager->closeDocument(tab);

        QCOMPARE(spy.count(), 1);
        QCOMPARE(spy.takeFirst().at(0).value<DocumentTab *>(), tab);
    }
};

QTEST_MAIN(TestDocumentManager)
#include "test_DocumentManager.moc"
