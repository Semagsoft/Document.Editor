#pragma once

#include <QObject>
#include <QString>
#include <QStringList>

class DocumentManager;
class DocumentEditor;
class QTemporaryDir;

class DocumentService : public QObject
{
    Q_OBJECT

public:
    explicit DocumentService(DocumentManager *docManager,
                             QWidget *parentWidget);
    ~DocumentService() override;

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

    static QString detectArchiver();
    static QString detectCompressor();

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
    // Emitted after a successful revert; receivers can drop per-editor state
    // such as the crash-recovery file for the reverted document.
    void documentReverted(DocumentEditor *editor);

private:
    void runArchiveExportAsync(const QString &src, const QString &path,
                               QTemporaryDir *tempDir);
    void runArchiveImportAsync(const QString &path);
    void runSoundExportAsync(const QString &path, const QString &text);

    DocumentManager *m_docManager;
    QWidget *m_parentWidget;
    // Temporary directories that back opened documents must outlive the
    // editor tabs; deleted when the service is destroyed (application exit).
    QList<QTemporaryDir*> m_archiveTempDirs;
};
