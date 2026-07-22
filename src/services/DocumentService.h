#pragma once

#include <QObject>
#include <QString>
#include <QStringList>

class DocumentManager;
class DocumentEditor;
class Settings;
class StatusBarManager;

class DocumentService : public QObject
{
    Q_OBJECT

public:
    explicit DocumentService(DocumentManager *docManager,
                             Settings *settings,
                             StatusBarManager *statusBar,
                             QWidget *parentWidget);

    void newDocument();
    void openDocument(const QString &path = QString());
    void importFtp();
    void importArchive();
    void importImage();
    void exportWordpress();
    void exportEmail();
    void exportFtp();
    void exportPdf();
    void exportArchive();
    void exportImage();
    void exportSound();
    void printDocument();
    void pageSetup();
    void revertDocument();

    static void embedImageInDocument(DocumentEditor *editor, const QString &path,
                                     qreal width = 0, qreal height = 0);

    static void insertTableInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertImageInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertShapeInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertChartInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertLinkInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertSymbolInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertHorizontalLineInteractive(DocumentEditor *editor);
    static void insertDateInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertTimeInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertVideoInteractive(QWidget *parent, DocumentEditor *editor);
    static void insertHeader(DocumentEditor *editor);
    static void insertFooter(DocumentEditor *editor);

signals:
    void documentOpened(const QString &path);
    void statusMessage(const QString &message, int timeout);

private:
    static bool findTool(const QString &name, const QStringList &args = {});
    static QString findArchiver();
    static QString findCompressor();

    DocumentManager *m_docManager;
    Settings *m_settings;
    StatusBarManager *m_statusBarManager;
    QWidget *m_parentWidget;
};
