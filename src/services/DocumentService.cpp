#include "services/DocumentService.h"
#include "converters/FtpClient.h"
#include "editor/DocumentEditor.h"
#include "editor/DocumentManager.h"

#include <QApplication>
#include <QColorDialog>
#include <QDate>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
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
#include <QUrl>

#include "dialogs/InsertChartDialog.h"
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

bool DocumentService::findTool(const QString &name, const QStringList &args)
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

QStringList DocumentService::listArchiveContents(const QString &archiver,
                                                 const QString &path)
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

QString DocumentService::findArchiver()
{
    if (findTool(QStringLiteral("7z")))
        return QStringLiteral("7z");
    if (findTool(QStringLiteral("unzip")))
        return QStringLiteral("unzip");
    return {};
}

QString DocumentService::findCompressor()
{
    if (findTool(QStringLiteral("7z")))
        return QStringLiteral("7z");
    if (findTool(QStringLiteral("zip"), { QStringLiteral("--version") }))
        return QStringLiteral("zip");
    return {};
}

void DocumentService::embedImageInDocument(DocumentEditor *editor, const QString &path,
                                           qreal width, qreal height)
{
    QImage image(path);
    if (image.isNull())
        return;
    QString name = QStringLiteral("embed_%1").arg(QFileInfo(path).fileName());
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
    editor->textCursor().insertImage(fmt);
}

void DocumentService::insertTableInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertTableDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted)
        editor->textCursor().insertTable(dlg.rows(), dlg.columns());
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
        if (!shape.isNull())
            editor->textCursor().insertImage(shape.toImage());
    }
}

void DocumentService::insertChartInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertChartDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QPixmap chart = dlg.generateChart(dlg.chartType());
        if (!chart.isNull())
            editor->textCursor().insertImage(chart.toImage());
    }
}

void DocumentService::insertLinkInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertLinkDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QTextCursor cursor = editor->textCursor();
        cursor.insertHtml(QStringLiteral("<a href=\"%1\">%2</a>")
            .arg(dlg.url().toHtmlEscaped(), dlg.displayText().toHtmlEscaped()));
    }
}

void DocumentService::insertSymbolInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertSymbolDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted)
        editor->textCursor().insertText(QString(dlg.selectedSymbol()));
}

void DocumentService::insertHorizontalLineInteractive(DocumentEditor *editor)
{
    editor->textCursor().insertHtml(QStringLiteral("<hr>"));
}

void DocumentService::insertDateInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertDateDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted)
        editor->textCursor().insertText(QDate::currentDate().toString(dlg.dateFormat()));
}

void DocumentService::insertTimeInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertTimeDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted)
        editor->textCursor().insertText(QTime::currentTime().toString(dlg.timeFormat()));
}

void DocumentService::insertVideoInteractive(QWidget *parent, DocumentEditor *editor)
{
    InsertVideoDialog dlg(parent);
    if (dlg.exec() == QDialog::Accepted) {
        QTextCursor cursor = editor->textCursor();
        cursor.insertHtml(QStringLiteral("<a href=\"%1\">%1</a>").arg(dlg.videoPath().toHtmlEscaped()));
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

    if (m_docManager->openDocument(filePath))
        emit documentOpened(filePath);
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
    QString archiver = findArchiver();
    if (archiver.isEmpty()) {
        QMessageBox::warning(m_parentWidget, tr("Archive Import"),
            tr("No archive tool found.\n\n"
               "Install 7-Zip (7z) or unzip:\n"
               "  sudo apt install 7zip unzip   (Debian/Ubuntu)\n"
               "  sudo dnf install 7zip unzip   (Fedora)\n"
               "  brew install 7zip unzip        (macOS)"));
        return;
    }

    // Reject archives with entries that escape the extraction directory
    // (zip-slip). 7z does not sanitize ".." entries, so validate before extracting.
    const QStringList listing = listArchiveContents(archiver, path);
    if (listing.isEmpty()) {
        QMessageBox::warning(m_parentWidget, tr("Archive Import"),
            tr("Could not read the archive contents."));
        return;
    }
    for (const QString &entry : listing) {
        const QString clean = QDir::cleanPath(entry);
        if (QDir::isAbsolutePath(clean) || clean == QStringLiteral("..")
            || clean.startsWith(QStringLiteral("../"))) {
            QMessageBox::warning(m_parentWidget, tr("Archive Import"),
                tr("Archive contains unsafe path entries and was not extracted:\n%1")
                    .arg(entry));
            return;
        }
    }

    QTemporaryDir tempDir;
    if (!tempDir.isValid())
        return;
    QProcess proc;
    proc.setWorkingDirectory(tempDir.path());
    if (archiver == QStringLiteral("7z"))
        proc.start(QStringLiteral("7z"), { QStringLiteral("x"), path, QStringLiteral("-y") });
    else
        proc.start(QStringLiteral("unzip"), { path });
    waitForProcess(proc, 30000);
    if (proc.exitCode() != 0) {
        QMessageBox::warning(m_parentWidget, tr("Archive Import"),
            tr("Failed to extract archive."));
        return;
    }
    QDir dir(tempDir.path());
    QStringList filters = { QStringLiteral("*.xaml"), QStringLiteral("*.html"),
        QStringLiteral("*.htm"), QStringLiteral("*.rtf"),
        QStringLiteral("*.txt") };
    QStringList files = dir.entryList(filters, QDir::Files, QDir::Name);
    if (files.isEmpty()) {
        QMessageBox::information(m_parentWidget, tr("Archive Import"),
            tr("No supported documents found in the archive."));
        return;
    }
    for (const QString &f : files)
        openDocument(dir.absoluteFilePath(f));
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
    if (editor->saveToFile(path))
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
        QPageLayout layout(QPageSize(QSizeF(pageW, pageH), QPageSize::Point),
                           QPageLayout::Portrait,
                           editor->pageMargins(),
                           QPageLayout::Point);
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
    QTemporaryDir tmpDir;
    if (src.isEmpty()) {
        if (!tmpDir.isValid())
            return;
        QString tmpFile = tmpDir.path() + QStringLiteral("/document.xaml");
        editor->saveToFile(tmpFile);
        src = tmpFile;
    }
    QString compressor = findCompressor();
    if (compressor.isEmpty()) {
        QMessageBox::warning(m_parentWidget, tr("Archive Export"),
            tr("No archive tool found.\n\n"
               "Install 7-Zip (7z) or zip:\n"
               "  sudo apt install 7zip zip   (Debian/Ubuntu)\n"
               "  sudo dnf install 7zip zip   (Fedora)\n"
               "  brew install 7zip zip        (macOS)"));
        return;
    }
    QProcess proc;
    if (compressor == QStringLiteral("7z")) {
        proc.setWorkingDirectory(QFileInfo(src).absolutePath());
        proc.start(QStringLiteral("7z"), { QStringLiteral("a"), path, QFileInfo(src).fileName() });
    } else {
        QStringList args;
        args << path << QFileInfo(src).fileName();
        proc.setWorkingDirectory(QFileInfo(src).absolutePath());
        proc.start(QStringLiteral("zip"), args);
    }
    waitForProcess(proc, 30000);
    if (proc.exitCode() == 0)
        emit statusMessage(tr("Archive saved"), 3000);
    else
        QMessageBox::warning(m_parentWidget, tr("Archive Export"),
            tr("Failed to create archive."));
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
    QPixmap pixmap(pageSize.toSize());
    pixmap.fill(Qt::white);

    // Inset the content by the configured page margins (relative to whatever the
    // document's root frame already uses) so the image matches print/PDF output.
    QPainter painter(&pixmap);
    const QMarginsF margins = editor->pageMargins();
    const QTextFrameFormat frameFmt = doc->rootFrame()->frameFormat();
    painter.translate(margins.left() - frameFmt.leftMargin(),
                      margins.top() - frameFmt.topMargin());
    doc->drawContents(&painter);
    painter.end();

    if (!pixmap.save(path))
        QMessageBox::warning(m_parentWidget, tr("Export Error"), tr("Could not save image."));
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
    if (findTool(QStringLiteral("espeak"), { QStringLiteral("--version") })) {
        QProcess proc;
        proc.start(QStringLiteral("espeak"), { QStringLiteral("-w"), path, text });
        waitForProcess(proc, 60000);
        if (proc.exitCode() == 0) {
            emit statusMessage(tr("Sound exported as WAV"), 3000);
            return;
        }
    }
    if (findTool(QStringLiteral("spd-say"), { QStringLiteral("--version") })) {
        QProcess proc;
        proc.start(QStringLiteral("spd-say"), { QStringLiteral("-w"), path, QStringLiteral("-o"), text });
        waitForProcess(proc, 60000);
        if (proc.exitCode() == 0) {
            emit statusMessage(tr("Sound exported as WAV"), 3000);
            return;
        }
    }
    QMessageBox::warning(m_parentWidget, tr("Export to Sound"),
        tr("No speech synthesis engine found.\n\n"
           "Install espeak or speech-dispatcher:\n"
           "  sudo apt install espeak          (Debian/Ubuntu)\n"
           "  sudo dnf install espeak          (Fedora)\n"
           "  brew install espeak              (macOS)"));
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
    e->loadFromFile(e->documentName());
}
