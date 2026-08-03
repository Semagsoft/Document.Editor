#include <QTest>
#include <QMainWindow>
#include <QMdiArea>
#include <QMdiSubWindow>
#include <QTextDocument>
#include <QTemporaryFile>
#include <QFile>
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
};

QTEST_MAIN(TestDocumentManager)
#include "test_DocumentManager.moc"
