#include "services/DocumentService.h"
#include "converters/FtpClient.h"
#include "editor/DocumentEditor.h"
#include "editor/DocumentManager.h"
#include "editor/DocumentTab.h"

#include <QApplication>
#include <QColorDialog>
#include <QCryptographicHash>
#include <QDate>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QFutureWatcher>
#include <QImage>
#include <QInputDialog>
#include <QMessageBox>
#include <QPageSetupDialog>
#include <QPainter>
#include <QPixmap>
#include <QPrintDialog>
#include <QPrinter>
#include <QProcess>
#include <QStandardPaths>
#include <QTextDocument>
#include <QTextFrame>
#include <QTextImageFormat>
#include <QTextStream>
#include <QTime>
#include <QTemporaryDir>
#include <QUuid>
#include <QUrl>

#include <QtConcurrent>

#include "dialogs/InsertChartDialog.h"
#include "dialogs/ChartData.h"
#include "dialogs/InsertDateDialog.h"
#include "dialogs/InsertImageDialog.h"
#include "dialogs/InsertLineDialog.h"
#include "dialogs/InsertLinkDialog.h"
#include "dialogs/InsertShapeDialog.h"
#include "dialogs/InsertSymbolDialog.h"
#include "dialogs/InsertTableDialog.h"
#include "dialogs/InsertTimeDialog.h"
#include "dialogs/InsertVideoDialog.h"

static void waitForProcess(QProcess &proc, int timeoutMs)
{
    if (proc.waitForFinished(timeoutMs))
        return;
    proc.kill();
    proc.waitForFinished(1000);
}

// Tool probing and subprocess execution run on a background thread (see the
// *Async helpers below) so the GUI thread never blocks on the 2-60s waits.

static bool findTool(const QString &name, const QStringList &args = {})
{
    QProcess proc;
    proc.start(name, args.isEmpty() ? QStringList{ QStringLiteral("--version") } : args);
    // waitForStarted fails if the executable does not exist / cannot be launched;
    // exitCode() alone would wrongly stay 0 for a never-started process.
    if (!proc.waitForStarted(2000))
        return false;
    waitForProcess(proc, 3000);
    return proc.exitStatus() == QProcess::NormalExit && proc.exitCode() == 0;
}

// 7-Zip and UnZip do not accept "--version": 7z exits 7 and unzip exits 10,
// so probe with their documented info flags instead.
static QString findArchiver()
{
    if (findTool(QStringLiteral("7z"), { QStringLiteral("i") }))
        return QStringLiteral("7z");
    if (findTool(QStringLiteral("unzip"), { QStringLiteral("-v") }))
        return QStringLiteral("unzip");
    return {};
}

static QString findCompressor()
{
    if (findTool(QStringLiteral("7z"), { QStringLiteral("i") }))
        return QStringLiteral("7z");
    if (findTool(QStringLiteral("zip"), { QStringLiteral("--version") }))
        return QStringLiteral("zip");
    return {};
}

QString DocumentService::detectArchiver()
{
    return findArchiver();
}

QString DocumentService::detectCompressor()
{
    return findCompressor();
}

static QStringList listArchiveContents(const QString &archiver, const QString &path)
{
    QStringList names;
    QProcess proc;
    if (archiver == QStringLiteral("7z"))
        proc.start(QStringLiteral("7z"), { QStringLiteral("l"), QStringLiteral("-slt"), path });
    else
        proc.start(QStringLiteral("unzip"), { QStringLiteral("-Z1"), path });
    if (!proc.waitForStarted(2000))
        return names;
    waitForProcess(proc, 15000);
    if (proc.exitStatus() != QProcess::NormalExit || proc.exitCode() != 0)
        return names;
    const QStringList lines =
        QString::fromUtf8(proc.readAllStandardOutput()).split(QLatin1Char('\n'));
    if (archiver == QStringLiteral("7z")) {
        for (const QString &line : lines) {
            if (line.startsWith(QLatin1String("Path = ")))
                names << line.mid(7).trimmed();
        }
    } else {
        names = lines;
    }
    names.removeAll(QString());
    return names;
}

struct ArchiveExportResult { bool toolMissing = false; bool success = false; };

static ArchiveExportResult runArchiveExport(const QString &src, const QString &path)
{
    ArchiveExportResult res;
    const QString compressor = findCompressor();
    if (compressor.isEmpty()) {
        res.toolMissing = true;
        return res;
    }
    QProcess proc;
    proc.setWorkingDirectory(QFileInfo(src).absolutePath());
    if (compressor == QStringLiteral("7z"))
        proc.start(QStringLiteral("7z"), { QStringLiteral("a"), path, QFileInfo(src).fileName() });
    else
        proc.start(QStringLiteral("zip"), { path, QFileInfo(src).fileName() });
    waitForProcess(proc, 30000);
    res.success = proc.exitStatus() == QProcess::NormalExit && proc.exitCode() == 0;
    return res;
}

struct ArchiveImportProbe {
    bool toolMissing = false;
    bool listFailed = false;
    bool unsafe = false;
    QString unsafeEntry;
    QString archiver;
    QStringList listing;
};

static ArchiveImportProbe probeArchiveImport(const QString &path)
{
    ArchiveImportProbe probe;
    const QString archiver = findArchiver();
    if (archiver.isEmpty()) {
        probe.toolMissing = true;
        return probe;
    }
    probe.archiver = archiver;
    // Reject archives with entries that escape the extraction directory
    // (zip-slip). 7z does not sanitize ".." entries, so validate before extracting.
    const QStringList listing = listArchiveContents(archiver, path);
    if (listing.isEmpty()) {
        probe.listFailed = true;
        return probe;
    }
    for (const QString &entry : listing) {
        const QString clean = QDir::cleanPath(entry);
        if (QDir::isAbsolutePath(clean) || clean == QStringLiteral("..")
            || clean.startsWith(QStringLiteral("../"))) {
            probe.unsafe = true;
            probe.unsafeEntry = entry;
            return probe;
        }
    }
    probe.listing = listing;
    return probe;
}

struct ArchiveImportOutcome { bool extractFailed = false; };

static ArchiveImportOutcome runArchiveImport(const QString &archiver,
                                             const QString &path,
                                             const QString &destDir)
{
    ArchiveImportOutcome outcome;
    QProcess proc;
    proc.setWorkingDirectory(destDir);
    if (archiver == QStringLiteral("7z"))
        proc.start(QStringLiteral("7z"), { QStringLiteral("x"), path, QStringLiteral("-y") });
    else
        proc.start(QStringLiteral("unzip"), { path });
    waitForProcess(proc, 30000);
    outcome.extractFailed = proc.exitStatus() != QProcess::NormalExit || proc.exitCode() != 0;
    return outcome;
}

struct SoundExportResult { bool success = false; };

static SoundExportResult runSoundExport(const QString &path, const QString &text)
{
    if (findTool(QStringLiteral("espeak"), { QStringLiteral("--version") })) {
        QProcess proc;
        proc.start(QStringLiteral("espeak"), { QStringLiteral("-w"), path, text });
        waitForProcess(proc, 60000);
        if (proc.exitStatus() == QProcess::NormalExit && proc.exitCode() == 0)
            return { true };
    }
    if (findTool(QStringLiteral("spd-say"), { QStringLiteral("--version") })) {
        QProcess proc;
        proc.start(QStringLiteral("spd-say"),
                   { QStringLiteral("-w"), path, QStringLiteral("-o"), text });
        waitForProcess(proc, 60000);
        if (proc.exitStatus() == QProcess::NormalExit && proc.exitCode() == 0)
            return { true };
    }
    return { false };
}

void DocumentService::embedImageInDocument(DocumentEditor *editor, const QString &path,
                                           qreal width, qreal height)
{
    QImage image(path);
    if (image.isNull())
        return;
    // Key the resource by the full path so two images that share a basename
    // (e.g. different folders) do not overwrite each other's resource, while
    // embedding the same file twice reuses the already-stored image.
    const QString name = QStringLiteral("embed_%1")
        .arg(QString::fromLatin1(QCryptographicHash::hash(path.toUtf8(),
                                                          QCryptographicHash::Sha1)
                                     .toHex()));
    editor->document()->addResource(QTextDocument::ImageResource, QUrl(name), image);
    QTextImageFormat fmt;
    fmt.setName(name);
    if (width > 0)
        fmt.setWidth(width);
    if (height > 0)
        fmt.setHeight(height);
    if (width <= 0 && height <= 0) {
        fmt.setWidth(image.width());
        fmt.setHeight(image.height());
    }
    QTextCursor cursor = editor->textCursor();
    cursor.insertImage(fmt);
    editor->setTextCursor(cursor);
}

void DocumentService::insertTableInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertTableDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QTextCursor cursor = editor->textCursor();
        cursor.insertTable(dlg.rows(), dlg.columns());
        editor->setTextCursor(cursor);
    }
}

void DocumentService::insertImageInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertImageDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted)
        embedImageInDocument(editor, dlg.imagePath(), dlg.width(), dlg.height());
}

void DocumentService::insertShapeInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertShapeDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QPixmap shape = dlg.generateShape(dlg.shapeName());
        if (!shape.isNull()) {
            QTextCursor cursor = editor->textCursor();
            cursor.insertImage(shape.toImage());
            editor->setTextCursor(cursor);
        }
    }
}

void DocumentService::insertChartInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertChartDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        ChartData data = dlg.chartData();
        QPixmap chart = dlg.generateChart(data.type);
        if (!chart.isNull()) {
            // Embed the rendered chart as an image resource so it actually
            // displays. A JSON-as-ImageResource placeholder renders blank.
            const QImage image = chart.toImage();
            const QString resourceName = QStringLiteral("chart_%1")
                .arg(QUuid::createUuid().toString(QUuid::WithoutBraces));
            editor->document()->addResource(QTextDocument::ImageResource,
                QUrl(resourceName), image);

            // Insert image referencing the resource; store type in image format name for editing
            QTextImageFormat fmt;
            fmt.setName(resourceName);
            fmt.setWidth(chart.width());
            fmt.setHeight(chart.height());
            fmt.setProperty(QTextFormat::ImageName, resourceName);
            fmt.setProperty(QTextFormat::UserProperty, data.type);
            QTextCursor cursor = editor->textCursor();
            cursor.insertImage(fmt);
            editor->setTextCursor(cursor);
        }
    }
}

void DocumentService::insertLinkInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertLinkDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        // toHtmlEscaped() does not escape quotes, so do that explicitly to
        // keep the URL from breaking out of the href attribute.
        auto escapeQuotes = [](const QString& in) {
            return QString(in).replace(QLatin1Char('"'), QStringLiteral("&quot;"));
        };
        QTextCursor cursor = editor->textCursor();
        cursor.insertHtml(QStringLiteral("<a href=\"%1\">%2</a>")
            .arg(escapeQuotes(dlg.url().trimmed().toHtmlEscaped()),
                 escapeQuotes(dlg.displayText().trimmed().toHtmlEscaped())));
        editor->setTextCursor(cursor);
    }
}

void DocumentService::insertSymbolInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertSymbolDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QTextCursor cursor = editor->textCursor();
        cursor.insertText(QString(dlg.selectedSymbol()));
        editor->setTextCursor(cursor);
    }
}

void DocumentService::insertHorizontalLineInteractive(DocumentEditor *editor)
{
    QTextCursor cursor = editor->textCursor();
    cursor.insertHtml(QStringLiteral("<hr>"));
    editor->setTextCursor(cursor);
}

void DocumentService::insertDateInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertDateDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QTextCursor cursor = editor->textCursor();
        cursor.insertText(QDate::currentDate().toString(dlg.dateFormat()));
        editor->setTextCursor(cursor);
    }
}

void DocumentService::insertTimeInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertTimeDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QTextCursor cursor = editor->textCursor();
        cursor.insertText(QTime::currentTime().toString(dlg.timeFormat()));
        editor->setTextCursor(cursor);
    }
}

void DocumentService::insertVideoInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertVideoDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        // toHtmlEscaped() does not escape quotes; the URL goes inside an href
        // attribute, so escape it explicitly.
        const QString url = dlg.videoPath().trimmed()
            .toHtmlEscaped()
            .replace(QLatin1Char('"'), QStringLiteral("&quot;"));
        QTextCursor cursor = editor->textCursor();
        cursor.insertHtml(QStringLiteral("<a href=\"%1\">%1</a>").arg(url));
        editor->setTextCursor(cursor);
    }
}

void DocumentService::insertHeader(DocumentEditor *editor)
{
    QTextCursor cursor = editor->textCursor();
    cursor.movePosition(QTextCursor::Start);
    cursor.insertHtml(QStringLiteral(
        "<div style=\"border-bottom: 2px solid #444; padding-bottom: 6px; "
        "margin-bottom: 12px; font-size: 10pt; color: #666;\">Header</div>"));
}

void DocumentService::insertFooter(DocumentEditor *editor)
{
    QTextCursor cursor = editor->textCursor();
    cursor.movePosition(QTextCursor::End);
    cursor.insertHtml(QStringLiteral(
        "<div style=\"border-top: 2px solid #444; padding-top: 6px; "
        "margin-top: 12px; font-size: 10pt; color: #666;\">Footer</div>"));
}

DocumentService::DocumentService(DocumentManager *docManager,
                                 QWidget *parentWidget)
    : QObject(parentWidget)
    , m_docManager(docManager)
    , m_parentWidget(parentWidget)
{
}

void DocumentService::newDocument()
{
    m_docManager->createDocument();
    emit documentOpened(QString());
}

void DocumentService::openDocument(const QString &path)
{
    QString filePath = path;
    if (filePath.isEmpty()) {
        filePath = QFileDialog::getOpenFileName(m_parentWidget, tr("Open Document"),
            QString(),
            tr("XAML Document (*.xaml *.dexml);;"
               "Word Document (*.docx);;"
               "HTML Document (*.html *.htm);;"
               "Rich Text Format (*.rtf);;"
               "Text File (*.txt);;"
               "All Files (*)"));
    }
    if (filePath.isEmpty())
        return;

    if (m_docManager->openDocument(filePath)) {
        if (DocumentEditor *e = m_docManager->activeEditor(); e && e->isReadOnlyFile())
            emit statusMessage(tr("Opened read-only — save with File > Save As"), 5000);
        emit documentOpened(filePath);
    }
}

void DocumentService::importFtp()
{
    bool ok = false;
    QString url = QInputDialog::getText(m_parentWidget, tr("FTP Import"),
        tr("Remote URL:"), QLineEdit::Normal, QStringLiteral("ftp://"), &ok);
    if (!ok || url.isEmpty())
        return;
    QString localPath = QFileDialog::getSaveFileName(m_parentWidget, tr("Save As"),
        QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation),
        tr("All Files (*)"));
    if (localPath.isEmpty())
        return;
    auto *ftp = new FtpClient(this);
    connect(ftp, &FtpClient::downloadCompleted, ftp, [this, ftp, localPath]() {
        openDocument(localPath);
        ftp->deleteLater();
    });
    connect(ftp, &FtpClient::transferFailed, ftp, [this, ftp](const QString &err) {
        QMessageBox::warning(m_parentWidget, tr("FTP Error"), err);
        ftp->deleteLater();
    });
    ftp->download(url, localPath);
}

void DocumentService::importArchive()
{
    QString path = QFileDialog::getOpenFileName(m_parentWidget, tr("Import Archive"),
        QString(), tr("Archives (*.zip *.tar.gz *.tar);;All Files (*)"));
    if (path.isEmpty())
        return;
    runArchiveImportAsync(path);
}

void DocumentService::runArchiveImportAsync(const QString &path)
{
    emit statusMessage(tr("Reading archive..."), 5000);
    auto *probeWatcher = new QFutureWatcher<ArchiveImportProbe>(this);
    connect(probeWatcher, &QFutureWatcher<ArchiveImportProbe>::finished, this,
            [this, probeWatcher, path]() {
        const ArchiveImportProbe probe = probeWatcher->result();
        probeWatcher->deleteLater();
        if (probe.toolMissing) {
            QMessageBox::warning(m_parentWidget, tr("Archive Import"),
                tr("No archive tool found.\n\n"
                   "Install 7-Zip (7z) or unzip:\n"
                   "  sudo apt install 7zip unzip   (Debian/Ubuntu)\n"
                   "  sudo dnf install 7zip unzip   (Fedora)\n"
                   "  brew install 7zip unzip        (macOS)"));
            return;
        }
        if (probe.unsafe) {
            QMessageBox::warning(m_parentWidget, tr("Archive Import"),
                tr("Archive contains unsafe path entries and was not extracted:\n%1")
                    .arg(probe.unsafeEntry));
            return;
        }
        if (probe.listFailed) {
            QMessageBox::warning(m_parentWidget, tr("Archive Import"),
                tr("Could not read the archive contents."));
            return;
        }

        // Extract to a temp dir kept alive until the documents are opened.
        // The watcher lambda captures a local pointer so concurrent imports
        // can't delete each other's temp dir.
        auto *tempDir = new QTemporaryDir();
        if (!tempDir->isValid()) {
            delete tempDir;
            QMessageBox::warning(m_parentWidget, tr("Archive Import"),
                tr("Could not create a temporary directory."));
            return;
        }
        emit statusMessage(tr("Extracting archive..."), 5000);
        auto *extractWatcher = new QFutureWatcher<ArchiveImportOutcome>(this);
        connect(extractWatcher, &QFutureWatcher<ArchiveImportOutcome>::finished, this,
                [this, extractWatcher, probe, tempDir]() {
            const ArchiveImportOutcome outcome = extractWatcher->result();
            extractWatcher->deleteLater();
            if (outcome.extractFailed || !tempDir) {
                delete tempDir;
                QMessageBox::warning(m_parentWidget, tr("Archive Import"),
                    tr("Failed to extract archive."));
                return;
            }
            QDir dir(tempDir->path());
            const QStringList filters = { QStringLiteral("*.xaml"), QStringLiteral("*.html"),
                QStringLiteral("*.htm"), QStringLiteral("*.rtf"),
                QStringLiteral("*.txt") };
            const QStringList files = dir.entryList(filters, QDir::Files, QDir::Name);
            if (files.isEmpty()) {
                delete tempDir;
                QMessageBox::information(m_parentWidget, tr("Archive Import"),
                    tr("No supported documents found in the archive."));
                return;
            }
            for (const QString &f : files)
                openDocument(dir.absoluteFilePath(f));
            delete tempDir;
        });
        extractWatcher->setFuture(QtConcurrent::run([probe, path,
                                                     dest = tempDir->path()]() {
            return runArchiveImport(probe.archiver, path, dest);
        }));
    });
    probeWatcher->setFuture(QtConcurrent::run([path]() {
        return probeArchiveImport(path);
    }));
}

void DocumentService::importImage()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString path = QFileDialog::getOpenFileName(m_parentWidget, tr("Insert Image"),
        QString(), tr("Images (*.png *.jpg *.jpeg *.gif *.bmp *.svg)"));
    if (!path.isEmpty())
        embedImageInDocument(editor, path);
}

void DocumentService::exportWordpress()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString path = QFileDialog::getSaveFileName(m_parentWidget, tr("Export to WordPress"),
        QString(), tr("HTML (*.html)"));
    if (path.isEmpty())
        return;
    if (editor->saveToFile(path, false))
        emit statusMessage(tr("Exported as HTML — import into WordPress"), 5000);
    else
        QMessageBox::warning(m_parentWidget, tr("Export Error"), tr("Could not save HTML file."));
}

void DocumentService::exportEmail()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString path = QFileDialog::getSaveFileName(m_parentWidget, tr("Export to Email"),
        QString(), tr("EML (*.eml)"));
    if (path.isEmpty())
        return;
    QFile file(path);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QMessageBox::warning(m_parentWidget, tr("Export Error"), tr("Could not write file."));
        return;
    }
    QTextStream out(&file);
    out << QStringLiteral("From: user@example.com\n");
    out << QStringLiteral("To: recipient@example.com\n");
    out << QStringLiteral("Subject: Document\n");
    out << QStringLiteral("MIME-Version: 1.0\n");
    out << QStringLiteral("Content-Type: text/html; charset=utf-8\n\n");
    out << editor->toHtml();
    file.close();
    emit statusMessage(tr("Email exported as .eml"), 3000);
}

void DocumentService::exportFtp()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString localPath = editor->documentName();
    if (localPath.isEmpty()) {
        QMessageBox::information(m_parentWidget, tr("FTP Export"),
            tr("Save the document locally first."));
        return;
    }
    bool ok = false;
    QString url = QInputDialog::getText(m_parentWidget, tr("FTP Export"),
        tr("Remote URL:"), QLineEdit::Normal,
        QStringLiteral("ftp://") + QFileInfo(localPath).fileName(), &ok);
    if (!ok || url.isEmpty())
        return;
    auto *ftp = new FtpClient(this);
    connect(ftp, &FtpClient::uploadCompleted, ftp, [this, ftp]() {
        emit statusMessage(tr("FTP upload complete"), 3000);
        ftp->deleteLater();
    });
    connect(ftp, &FtpClient::transferFailed, ftp, [this, ftp](const QString &err) {
        QMessageBox::warning(m_parentWidget, tr("FTP Error"), err);
        ftp->deleteLater();
    });
    ftp->upload(localPath, url);
}

void DocumentService::exportPdf()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString path = QFileDialog::getSaveFileName(m_parentWidget, tr("Export as PDF"),
        QString(), tr("PDF (*.pdf)"));
    if (path.isEmpty())
        return;
    QPrinter printer(QPrinter::HighResolution);
    printer.setOutputFormat(QPrinter::PdfFormat);
    printer.setOutputFileName(path);

    // Match the printer layout to the document's page size and margins so the
    // rendered output lines up with what the user sees on screen.
    const qreal pageW = editor->pageWidth();
    const qreal pageH = editor->pageHeight();
    if (pageW > 0 && pageH > 0) {
        // The on-screen document is modeled in 96-DPI pixels; QPrinter works in
        // points (1/72"). Convert so an 8.5"x11" page exports at Letter size.
        const qreal pxToPt = 72.0 / 96.0;
        const QMarginsF m = editor->pageMargins();
        const QMarginsF marginsPt(m.left() * pxToPt, m.top() * pxToPt,
                                  m.right() * pxToPt, m.bottom() * pxToPt);
        const bool landscape = pageW > pageH;
        QPageLayout layout(
            QPageSize(QSizeF(pageW * pxToPt, pageH * pxToPt), QPageSize::Point),
            landscape ? QPageLayout::Landscape : QPageLayout::Portrait,
            marginsPt, QPageLayout::Point);
        printer.setPageLayout(layout);
    }

    // print() paginates the document across all pages (drawContents only paints
    // a single canvas, truncating multi-page documents to page one).
    editor->document()->print(&printer);
    emit statusMessage(tr("Saved as PDF"), 3000);
}

void DocumentService::exportArchive()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString path = QFileDialog::getSaveFileName(m_parentWidget, tr("Export to Archive"),
        QString(), tr("ZIP Archive (*.zip)"));
    if (path.isEmpty())
        return;
    if (!path.endsWith(QStringLiteral(".zip"), Qt::CaseInsensitive))
        path += QStringLiteral(".zip");
    QString src = editor->documentName();
    // Kept alive until the async archive job finishes; the lambda captures a
    // local pointer so concurrent jobs can't delete each other's temp dir.
    QTemporaryDir *tempDir = nullptr;
    if (src.isEmpty()) {
        // Materialize the unsaved document to a temp file owned by tempDir.
        tempDir = new QTemporaryDir();
        if (!tempDir->isValid()) {
            delete tempDir;
            QMessageBox::warning(m_parentWidget, tr("Archive Export"),
                tr("Could not create a temporary directory."));
            return;
        }
        QString tmpFile = tempDir->path() + QStringLiteral("/document.xaml");
        if (!editor->saveToFile(tmpFile, false)) {
            delete tempDir;
            QMessageBox::warning(m_parentWidget, tr("Archive Export"),
                tr("Could not write the temporary document."));
            return;
        }
        src = tmpFile;
    }
    runArchiveExportAsync(src, path, tempDir);
}

void DocumentService::runArchiveExportAsync(const QString &src, const QString &path,
                                            QTemporaryDir *tempDir)
{
    emit statusMessage(tr("Creating archive..."), 5000);
    auto *watcher = new QFutureWatcher<ArchiveExportResult>(this);
    connect(watcher, &QFutureWatcher<ArchiveExportResult>::finished, this,
            [this, watcher, tempDir]() {
        const ArchiveExportResult res = watcher->result();
        watcher->deleteLater();
        delete tempDir;
        if (res.toolMissing) {
            QMessageBox::warning(m_parentWidget, tr("Archive Export"),
                tr("No archive tool found.\n\n"
                   "Install 7-Zip (7z) or zip:\n"
                   "  sudo apt install 7zip zip   (Debian/Ubuntu)\n"
                   "  sudo dnf install 7zip zip   (Fedora)\n"
                   "  brew install 7zip zip        (macOS)"));
            return;
        }
        if (res.success)
            emit statusMessage(tr("Archive saved"), 3000);
        else
            QMessageBox::warning(m_parentWidget, tr("Archive Export"),
                tr("Failed to create archive."));
    });
    watcher->setFuture(QtConcurrent::run([src, path]() {
        return runArchiveExport(src, path);
    }));
}

void DocumentService::exportImage()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString path = QFileDialog::getSaveFileName(m_parentWidget, tr("Export as Image"),
        QString(), tr("PNG (*.png)"));
    if (path.isEmpty())
        return;
    QTextDocument *doc = editor->document();
    QSizeF pageSize = doc->pageSize();
    int pageCount = doc->pageCount();
    if (pageCount < 1)
        pageCount = 1;

    int pageW = qRound(pageSize.width());
    int pageH = qRound(pageSize.height());
    // Render every page stacked vertically so multi-page documents export
    // completely (drawContents only paints page one when the pixmap is
    // single-page sized).
    QPixmap pixmap(pageW, pageH * pageCount);
    pixmap.fill(Qt::white);

    // Inset the content by the configured page margins so the image matches
    // print/PDF output.
    QPainter painter(&pixmap);
    const QMarginsF margins = editor->pageMargins();
    const QTextFrameFormat frameFmt = doc->rootFrame()->frameFormat();
    painter.translate(margins.left() - frameFmt.leftMargin(),
                      margins.top() - frameFmt.topMargin());
    doc->drawContents(&painter);
    painter.end();

    if (!pixmap.save(path))
        QMessageBox::warning(m_parentWidget, tr("Export Error"), tr("Could not save image."));
    else
        emit statusMessage(tr("Exported %1 page(s) as image").arg(pageCount), 3000);
}

void DocumentService::exportSound()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QString path = QFileDialog::getSaveFileName(m_parentWidget, tr("Export to Sound"),
        QString(), tr("WAV (*.wav)"));
    if (path.isEmpty())
        return;
    QString text = editor->toPlainText().left(2000);
    if (text.isEmpty()) {
        QMessageBox::information(m_parentWidget, tr("Export to Sound"),
            tr("No text content to export."));
        return;
    }
    runSoundExportAsync(path, text);
}

void DocumentService::runSoundExportAsync(const QString &path, const QString &text)
{
    emit statusMessage(tr("Exporting to sound..."), 5000);
    auto *watcher = new QFutureWatcher<SoundExportResult>(this);
    connect(watcher, &QFutureWatcher<SoundExportResult>::finished, this,
            [this, watcher]() {
        const SoundExportResult res = watcher->result();
        watcher->deleteLater();
        if (res.success) {
            emit statusMessage(tr("Sound exported as WAV"), 3000);
            return;
        }
        QMessageBox::warning(m_parentWidget, tr("Export to Sound"),
            tr("No speech synthesis engine found.\n\n"
               "Install espeak or speech-dispatcher:\n"
               "  sudo apt install espeak          (Debian/Ubuntu)\n"
               "  sudo dnf install espeak          (Fedora)\n"
               "  brew install espeak              (macOS)"));
    });
    watcher->setFuture(QtConcurrent::run([path, text]() {
        return runSoundExport(path, text);
    }));
}

void DocumentService::printDocument()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QPrinter printer(QPrinter::HighResolution);
    QPrintDialog dlg(&printer, m_parentWidget);
    if (dlg.exec() == QDialog::Accepted)
        editor->document()->print(&printer);
}

void DocumentService::pageSetup()
{
    DocumentEditor *editor = m_docManager->activeEditor();
    if (!editor)
        return;
    QPrinter printer(QPrinter::HighResolution);
    printer.setPageSize(QPageSize(QPageSize::Letter));
    QPageSetupDialog dlg(&printer, m_parentWidget);
    if (dlg.exec() == QDialog::Accepted) {
        QPageSize pageSize = printer.pageLayout().pageSize();
        if (pageSize.isValid()) {
            QSize px = pageSize.sizePixels(96);
            editor->setPageWidth(px.width());
            editor->setPageHeight(px.height());
        }
        editor->setPageMargins(printer.pageLayout().marginsPixels(96));
        if (DocumentTab *tab = m_docManager->activeTab())
            tab->syncRuler();
    }
}

void DocumentService::revertDocument()
{
    DocumentEditor *e = m_docManager->activeEditor();
    if (!e || e->documentName().isEmpty())
        return;
    if (e->isModified()) {
        auto btn = QMessageBox::question(m_parentWidget, tr("Revert"),
            tr("Revert to the last saved version?\nAll changes will be lost."),
            QMessageBox::Yes | QMessageBox::No);
        if (btn != QMessageBox::Yes)
            return;
    }
    QString errorDetail;
    if (!e->loadFromFile(e->documentName(), &errorDetail)) {
        QString detail = errorDetail.isEmpty() ? tr("Unknown error.") : errorDetail;
        QMessageBox::warning(m_parentWidget, tr("Error"),
            tr("Could not revert file:\n%1\n\n%2").arg(e->documentName(), detail));
    }
}
