#pragma once

#include <QObject>
#include <QString>

class QMdiArea;
class QMdiSubWindow;
class DocumentTab;
class DocumentEditor;
class QWidget;

class DocumentManager : public QObject
{
    Q_OBJECT

public:
    explicit DocumentManager(QMdiArea *mdiArea, QWidget *mainWindowWidget,
                             QObject *parent = nullptr);

    DocumentTab *createDocument(const QString &title = QString());
    DocumentTab *openDocument(const QString &path);
    bool saveDocument(DocumentTab *tab = nullptr);
    bool saveDocumentAs(DocumentTab *tab = nullptr);
    bool saveDocumentCopy(DocumentTab *tab = nullptr);
    void saveAllDocuments();

    bool closeDocument(DocumentTab *tab = nullptr);
    void closeAllDocuments();
    void closeAllButCurrent();

    DocumentTab *activeTab() const;
    DocumentEditor *activeEditor() const;
    QList<DocumentTab *> allTabs() const;

    int tabCount() const;

signals:
    void activeTabChanged(DocumentTab *tab);
    void statusLineColumnChanged(int line, int col, int totalLines, int lineCount);
    void statusWordCountChanged(int wordCount);
    // Emitted when a tab is committed to closing (the save prompt, if any,
    // has been resolved) but before it is removed and destroyed. Receivers
    // can use this to discard per-tab state such as crash-recovery files.
    void tabAboutToClose(DocumentTab *tab);

protected:
    bool eventFilter(QObject *obj, QEvent *event) override;

private:
    void setupSubWindow(QMdiSubWindow *subWindow, DocumentTab *tab);
    void onSubWindowActivated(QMdiSubWindow *subWindow);

    QMdiArea *m_mdiArea = nullptr;
    QWidget *m_mainWindow = nullptr;
};
