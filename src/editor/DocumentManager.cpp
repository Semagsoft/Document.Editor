#include "DocumentManager.h"
#include "DocumentEditor.h"
#include "DocumentTab.h"

#include <QWidget>

#include <QApplication>
#include <QFileDialog>
#include <QFileInfo>
#include <QMdiArea>
#include <QMdiSubWindow>
#include <QMessageBox>
#include <QPointer>
#include <QStyle>

DocumentManager::DocumentManager(QMdiArea* mdiArea, QWidget* mainWindow,
    QObject* parent)
    : QObject(parent)
    , m_mdiArea(mdiArea)
    , m_mainWindow(mainWindow)
{
    connect(m_mdiArea, &QMdiArea::subWindowActivated,
        this, &DocumentManager::onSubWindowActivated);
}

DocumentTab* DocumentManager::createDocument(const QString& title)
{
    QString tabTitle = title.isEmpty() ? tr("Untitled") : title;
    DocumentTab* tab = new DocumentTab(tabTitle);

    QMdiSubWindow* subWindow = m_mdiArea->addSubWindow(tab);
    setupSubWindow(subWindow, tab);

    subWindow->show();
    m_mdiArea->setActiveSubWindow(subWindow);

    return tab;
}

DocumentTab* DocumentManager::openDocument(const QString& path)
{
    DocumentTab* tab = new DocumentTab();
    if (!tab->editor()->loadFromFile(path)) {
        QMessageBox::warning(m_mainWindow, tr("Error"),
            tr("Could not open file:\n%1").arg(path));
        delete tab;
        return nullptr;
    }

    tab->setDocumentName(path);

    QMdiSubWindow* subWindow = m_mdiArea->addSubWindow(tab);
    setupSubWindow(subWindow, tab);

    subWindow->show();
    m_mdiArea->setActiveSubWindow(subWindow);

    return tab;
}

bool DocumentManager::saveDocument(DocumentTab* tab)
{
    if (!tab)
        tab = activeTab();
    if (!tab)
        return false;

    DocumentEditor* editor = tab->editor();
    if (editor->documentName().isEmpty())
        return saveDocumentAs(tab);

    if (!editor->saveToFile(editor->documentName())) {
        QMessageBox::warning(m_mainWindow, tr("Error"),
            tr("Could not save file:\n%1").arg(editor->documentName()));
        return false;
    }

    emit documentSaved(tab);
    return true;
}

bool DocumentManager::saveDocumentAs(DocumentTab* tab)
{
    if (!tab)
        tab = activeTab();
    if (!tab)
        return false;

    QString path = QFileDialog::getSaveFileName(m_mainWindow, tr("Save As"),
        tab->documentName().isEmpty() ? QStringLiteral("Untitled.xaml") : tab->documentName(),
        tr("XAML Document (*.xaml);;"
           "HTML Document (*.html);;"
           "Rich Text Format (*.rtf);;"
           "Text File (*.txt);;"
           "All Files (*)"));

    if (path.isEmpty())
        return false;

    DocumentEditor* editor = tab->editor();
    if (!editor->saveToFile(path)) {
        QMessageBox::warning(m_mainWindow, tr("Error"),
            tr("Could not save file:\n%1").arg(path));
        return false;
    }

    tab->setDocumentName(path);
    emit documentSaved(tab);
    return true;
}

bool DocumentManager::saveDocumentCopy(DocumentTab* tab)
{
    if (!tab)
        tab = activeTab();
    if (!tab)
        return false;

    QString path = QFileDialog::getSaveFileName(m_mainWindow, tr("Save Copy"),
        tab->documentName(),
        tr("XAML Document (*.xaml);;"
           "HTML Document (*.html);;"
           "Rich Text Format (*.rtf);;"
           "Text File (*.txt);;"
           "All Files (*)"));

    if (path.isEmpty())
        return false;

    if (!tab->editor()->saveToFile(path))
        return false;

    emit documentSaved(tab);
    return true;
}

void DocumentManager::saveAllDocuments()
{
    for (DocumentTab* tab : allTabs())
        saveDocument(tab);
}

bool DocumentManager::closeDocument(DocumentTab* tab)
{
    if (!tab)
        tab = activeTab();
    if (!tab)
        return true;

    DocumentEditor* editor = tab->editor();
    if (editor->isModified()) {
        QMessageBox::StandardButton btn = QMessageBox::question(
            m_mainWindow, tr("Unsaved Changes"),
            tr("Save changes to '%1'?").arg(tab->tabTitle()),
            QMessageBox::Save | QMessageBox::Discard | QMessageBox::Cancel);

        if (btn == QMessageBox::Save)
            saveDocument(tab);
        else if (btn == QMessageBox::Cancel)
            return false;
    }

    QMdiSubWindow* sub = qobject_cast<QMdiSubWindow*>(tab->parentWidget());
    if (sub) {
        QString name = editor->documentName();
        m_mdiArea->removeSubWindow(sub);
        delete sub;
        emit documentClosed(name);
    }
    return true;
}

void DocumentManager::closeAllDocuments()
{
    QList<DocumentTab*> tabs = allTabs();
    for (DocumentTab* tab : tabs) {
        if (!tab)
            continue;
        if (!closeDocument(tab))
            return;
    }
    if (tabCount() == 0)
        emit allDocumentsClosed();
}

void DocumentManager::closeAllButCurrent()
{
    DocumentTab* current = activeTab();
    if (!current)
        return;

    QList<DocumentTab*> tabs = allTabs();
    for (DocumentTab* tab : tabs) {
        if (tab != current) {
            if (!closeDocument(tab))
                return;
        }
    }
}

DocumentTab* DocumentManager::activeTab() const
{
    QMdiSubWindow* active = m_mdiArea->activeSubWindow();
    if (!active)
        return nullptr;
    return qobject_cast<DocumentTab*>(active->widget());
}

DocumentEditor* DocumentManager::activeEditor() const
{
    DocumentTab* tab = activeTab();
    return tab ? tab->editor() : nullptr;
}

QList<DocumentTab*> DocumentManager::allTabs() const
{
    QList<DocumentTab*> tabs;
    for (QMdiSubWindow* sub : m_mdiArea->subWindowList())
        tabs.append(qobject_cast<DocumentTab*>(sub->widget()));
    return tabs;
}

int DocumentManager::tabCount() const
{
    return m_mdiArea->subWindowList().size();
}

void DocumentManager::setupSubWindow(QMdiSubWindow* subWindow, DocumentTab* tab)
{
    subWindow->setAttribute(Qt::WA_DeleteOnClose);
    subWindow->setWindowIcon(qApp->style()->standardIcon(QStyle::SP_FileIcon));

    QPointer<QMdiSubWindow> safeSubWindow(subWindow);
    connect(tab, &DocumentTab::titleChanged, this, [safeSubWindow](const QString& title) {
        if (safeSubWindow)
            safeSubWindow->setWindowTitle(title);
    });
}

void DocumentManager::onSubWindowActivated(QMdiSubWindow* subWindow)
{
    DocumentTab* tab = subWindow
        ? qobject_cast<DocumentTab*>(subWindow->widget())
        : nullptr;
    emit activeTabChanged(tab);

    if (tab && tab->editor()) {
        DocumentEditor* editor = tab->editor();
        emit statusLineColumnChanged(
            editor->selectedLineNumber(),
            editor->selectedColumnNumber(),
            editor->lineCount(),
            editor->columnCount());
        emit statusWordCountChanged(editor->wordCount());
    }
}
