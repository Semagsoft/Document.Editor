#include "DocumentEditor.h"
#include "converters/DocxConverter.h"
#include "converters/RtfConverter.h"
#include "converters/XamlConverter.h"
#include "services/SpellCheckHighlighter.h"
#include "utils/SpellChecker.h"
#include "utils/TextUtils.h"

#include <QAction>
#include <QApplication>
#include <QBuffer>
#include <QClipboard>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QFontDatabase>
#include <QKeyEvent>
#include <QMimeData>
#include <QRegularExpression>
#include <QSaveFile>
#include <QTextBlock>
#include <QTextCursor>
#include <QTextDocumentFragment>
#include <QTextFragment>
#include <QTextFrame>
#include <QTextList>
#include <QTextTable>

// Returns true if `path` is safe to open or write. Rejects:
//   * empty names
//   * any ".." path component (directory traversal)
//   * absolute paths that escape the allowed root (documents folder or the
//     current document's directory, whichever applies)
static bool isSafeFilePath(const QString& path, const QString& allowedRoot = QString())
{
    if (path.isEmpty())
        return false;

    QString cleaned = QDir::cleanPath(path);
    if (cleaned.isEmpty() || cleaned == QLatin1String("."))
        return false;

    // Reject any ".." component — this is the core traversal guard.
    const QStringList parts = cleaned.split(QLatin1Char('/'), Qt::SkipEmptyParts);
    for (const QString& part : parts) {
        if (part == QLatin1String(".."))
            return false;
    }

    // If an allowed root is given, absolute paths must stay within it.
    if (!allowedRoot.isEmpty() && QDir::isAbsolutePath(cleaned)) {
        QString root = QDir::cleanPath(allowedRoot);
        if (!cleaned.startsWith(root + QLatin1String("/")) && cleaned != root)
            return false;
    }

    return true;
}

DocumentEditor::DocumentEditor(QWidget* parent)
    : QTextEdit(parent)
{
    setAcceptRichText(true);
    setTabChangesFocus(false);
    setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    setLineWrapMode(QTextEdit::WidgetWidth);
    setWordWrapMode(QTextOption::WordWrap);

    constexpr QSizeF kLetterSize(816, 1056);
    QTextDocument* doc = document();
    doc->setPageSize(kLetterSize);
    doc->setDocumentMargin(0);

    // Default font
    m_baseFont = QFont(QStringLiteral("Segoe UI"), 12);
    setFont(m_baseFont);
    doc->setDefaultFont(m_baseFont);

    // Set default paragraph format: 1.15 line spacing
    QTextBlockFormat blockFmt;
    blockFmt.setLineHeight(115, QTextBlockFormat::ProportionalHeight);
    blockFmt.setTopMargin(0);
    blockFmt.setBottomMargin(0);
    QTextCursor cursor(doc);
    cursor.select(QTextCursor::Document);
    cursor.mergeBlockFormat(blockFmt);

    m_baseFontPointSize = m_baseFont.pointSizeF();

    m_statsTimer = new QTimer(this);
    m_statsTimer->setSingleShot(true);
    m_statsTimer->setInterval(200);
    connect(m_statsTimer, &QTimer::timeout, this, &DocumentEditor::updateStats);

    connectDocumentSignals();
    updateStats();
}

void DocumentEditor::connectDocumentSignals()
{
    connect(this, &QTextEdit::textChanged, this, [this]() {
        if (!m_fileChanged) {
            m_fileChanged = true;
            emit modifiedChanged(true);
        }
        scheduleStatsUpdate();
    });

    connect(this, &QTextEdit::cursorPositionChanged, this, [this]() {
        scheduleStatsUpdate();
        emit cursorPositionUpdated();
    });
}

void DocumentEditor::refreshStats()
{
    updateStats();
}

void DocumentEditor::scheduleStatsUpdate()
{
    if (!m_statsTimer->isActive())
        m_statsTimer->start();
}

// ============================================================
// Document State
// ============================================================

QString DocumentEditor::documentName() const { return m_documentName; }
void DocumentEditor::setDocumentName(const QString& name)
{
    m_documentName = name;
    QFileInfo fi(name);
    m_readOnlyFile = fi.exists() && !fi.isWritable();
    setReadOnly(m_readOnlyFile);
}

bool DocumentEditor::isModified() const { return m_fileChanged; }

void DocumentEditor::setModified(bool changed)
{
    if (m_fileChanged != changed) {
        m_fileChanged = changed;
        emit modifiedChanged(changed);
    }
}

bool DocumentEditor::isReadOnlyFile() const { return m_readOnlyFile; }

// ============================================================
// Page Setup
// ============================================================

QMarginsF DocumentEditor::pageMargins() const { return m_pageMargins; }

void DocumentEditor::setPageMargins(const QMarginsF& margins)
{
    m_pageMargins = margins;
    QTextFrameFormat frameFmt = document()->rootFrame()->frameFormat();
    frameFmt.setLeftMargin(margins.left());
    frameFmt.setRightMargin(margins.right());
    frameFmt.setTopMargin(margins.top());
    frameFmt.setBottomMargin(margins.bottom());
    document()->rootFrame()->setFrameFormat(frameFmt);
}

qreal DocumentEditor::pageWidth() const { return document()->pageSize().width(); }

void DocumentEditor::setPageWidth(qreal width)
{
    QSizeF size = document()->pageSize();
    size.setWidth(width);
    document()->setPageSize(size);
}

qreal DocumentEditor::pageHeight() const { return document()->pageSize().height(); }

void DocumentEditor::setPageHeight(qreal height)
{
    QSizeF size = document()->pageSize();
    size.setHeight(height);
    document()->setPageSize(size);
}

// ============================================================
// Zoom
// ============================================================

qreal DocumentEditor::zoomLevel() const { return m_zoomLevel; }

void DocumentEditor::setZoomLevel(qreal level)
{
    level = qBound(0.1, level, 5.0);
    const qreal factor = m_zoomLevel > 0 ? level / m_zoomLevel : 1.0;
    const bool scaleContent = !qFuzzyCompare(m_zoomLevel, level);
    m_zoomLevel = level;

    QFont zoomedFont = m_baseFont;
    zoomedFont.setPointSizeF(m_baseFontPointSize * m_zoomLevel);
    setFont(zoomedFont);

    // setFont() only rescales text that has no explicit char format. Scale
    // runs that carry an explicit font size too, so loaded/formatted documents
    // actually respond to zoom.
    if (!scaleContent)
        return;

    QTextDocument *doc = document();
    QTextCursor editCursor(doc);
    editCursor.beginEditBlock();

    // Walk every frame (root + table cells) so formatted runs inside tables
    // respond to zoom too. A rootFrame()-only iteration skips child frames.
    auto scaleFragments = [this, factor, doc](auto &&self, QTextFrame *frame) -> void {
        for (QTextFrame::iterator fit = frame->begin(); !fit.atEnd(); ++fit) {
            if (QTextFrame *child = fit.currentFrame()) {
                self(self, child);
                continue;
            }
            const QTextBlock block = fit.currentBlock();
            if (!block.isValid())
                continue;
            for (QTextBlock::iterator bit = block.begin(); !bit.atEnd(); ++bit) {
                const QTextFragment fragment = bit.fragment();
                if (!fragment.isValid() || fragment.charFormat().fontPointSize() <= 0)
                    continue;
                QTextCharFormat fmt = fragment.charFormat();
                fmt.setFontPointSize(fmt.fontPointSize() * factor);
                QTextCursor c(doc);
                c.setPosition(fragment.position());
                c.setPosition(fragment.position() + fragment.length(), QTextCursor::KeepAnchor);
                c.setCharFormat(fmt);
            }
        }
    };
    scaleFragments(scaleFragments, doc->rootFrame());

    editCursor.endEditBlock();
}

void DocumentEditor::setBaseFont(const QFont& font)
{
    m_baseFont = font;
    if (m_baseFont.pointSizeF() > 0)
        m_baseFontPointSize = m_baseFont.pointSizeF();
    document()->setDefaultFont(m_baseFont);
    setZoomLevel(m_zoomLevel);
}

void DocumentEditor::setBaseFontPointSize(qreal size)
{
    QFont font = m_baseFont;
    font.setPointSizeF(size);
    setBaseFont(font);
}

// ============================================================
// Statistics
// ============================================================

void DocumentEditor::updateStats()
{
    QTextCursor cursor = textCursor();
    QTextDocument* doc = document();

    // Line count and current line
    int blockNum = cursor.blockNumber() + 1;
    int totalBlocks = doc->blockCount();

    // Column
    int col = cursor.positionInBlock() + 1;
    int maxCol = col;
    QTextBlock block = cursor.block();
    if (block.isValid())
        // QTextBlock::length() counts the paragraph separator — the visible
        // column count is the plain-text width of the line.
        maxCol = block.text().length();

    m_selectedLine = blockNum;
    m_lineCount = totalBlocks;
    m_selectedColumn = col;
    m_columnCount = maxCol;

    // Word count
    QString text = doc->toPlainText();
    text = text.replace(QRegularExpression(QStringLiteral("[\\r\\n]+")), QStringLiteral(" "));
    text = text.trimmed();
    m_wordCount = 0;
    if (!text.isEmpty()) {
        QStringList words = text.split(QRegularExpression(QStringLiteral("\\s+")),
            Qt::SkipEmptyParts);
        m_wordCount = words.size();
    }
}

int DocumentEditor::lineCount() const { return m_lineCount; }
int DocumentEditor::columnCount() const { return m_columnCount; }
int DocumentEditor::wordCount() const { return m_wordCount; }
int DocumentEditor::selectedLineNumber() const { return m_selectedLine; }
int DocumentEditor::selectedColumnNumber() const { return m_selectedColumn; }

// ============================================================
// Find
// ============================================================

QTextCursor DocumentEditor::findWord(const QString& word, const QTextCursor& start)
{
    QTextCursor cursor = document()->find(word, start.isNull() ? textCursor() : start);
    if (!cursor.isNull()) {
        setTextCursor(cursor);
        ensureCursorVisible();
    }
    return cursor;
}

void DocumentEditor::goToLine(int lineNumber)
{
    if (lineNumber < 1)
        lineNumber = 1;
    QTextBlock block = document()->findBlockByNumber(lineNumber - 1);
    if (!block.isValid())
        block = document()->lastBlock();
    QTextCursor cursor(block);
    setTextCursor(cursor);
    ensureCursorVisible();
}

// ============================================================
// Formatting Toggles
// ============================================================

void DocumentEditor::toggleBold()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setFontWeight(cursor.charFormat().fontWeight() == QFont::Bold
            ? QFont::Normal
            : QFont::Bold);
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toggleItalic()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setFontItalic(!cursor.charFormat().fontItalic());
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toggleUnderline()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setFontUnderline(!cursor.charFormat().fontUnderline());
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toggleStrikethrough()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setFontStrikeOut(!cursor.charFormat().fontStrikeOut());
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toggleSubscript()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setVerticalAlignment(
        cursor.charFormat().verticalAlignment() == QTextCharFormat::AlignSubScript
            ? QTextCharFormat::AlignNormal
            : QTextCharFormat::AlignSubScript);
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toggleSuperscript()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setVerticalAlignment(
        cursor.charFormat().verticalAlignment() == QTextCharFormat::AlignSuperScript
            ? QTextCharFormat::AlignNormal
            : QTextCharFormat::AlignSuperScript);
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::clearFormatting()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    if (!cursor.hasSelection())
        cursor.select(QTextCursor::WordUnderCursor);
    fmt.setFont(QFont());
    fmt.setFontWeight(QFont::Normal);
    fmt.setFontItalic(false);
    fmt.setFontUnderline(false);
    fmt.setFontStrikeOut(false);
    fmt.setVerticalAlignment(QTextCharFormat::AlignNormal);
    cursor.setCharFormat(fmt);
    cursor.endEditBlock();
    setTextCursor(cursor);
}

// ============================================================
// Alignment
// ============================================================

void DocumentEditor::setParagraphAlignment(Qt::Alignment align)
{
    QTextBlockFormat fmt;
    fmt.setAlignment(align);
    QTextCursor cursor = textCursor();
    cursor.mergeBlockFormat(fmt);
    setTextCursor(cursor);
}

// ============================================================
// Lists
// ============================================================

void DocumentEditor::toggleBulletList()
{
    QTextCursor cursor = textCursor();
    QTextBlock block = cursor.block();
    QTextList* list = block.textList();

    if (list && list->format().style() == QTextListFormat::ListDisc) {
        list->remove(block);
        QTextBlockFormat bfmt;
        bfmt.setIndent(0);
        cursor.setBlockFormat(bfmt);
    } else {
        QTextListFormat listFmt;
        listFmt.setStyle(QTextListFormat::ListDisc);
        listFmt.setIndent(1);
        cursor.createList(listFmt);
    }
    setTextCursor(cursor);
}

void DocumentEditor::toggleNumberList()
{
    QTextCursor cursor = textCursor();
    QTextBlock block = cursor.block();
    QTextList* list = block.textList();

    if (list && list->format().style() == QTextListFormat::ListDecimal) {
        list->remove(block);
        QTextBlockFormat bfmt;
        bfmt.setIndent(0);
        cursor.setBlockFormat(bfmt);
    } else {
        QTextListFormat listFmt;
        listFmt.setStyle(QTextListFormat::ListDecimal);
        listFmt.setIndent(1);
        cursor.createList(listFmt);
    }
    setTextCursor(cursor);
}

// ============================================================
// Indent
// ============================================================

void DocumentEditor::indentMore()
{
    QTextCursor cursor = textCursor();
    QTextBlockFormat fmt;
    fmt.setIndent(cursor.blockFormat().indent() + 1);
    cursor.mergeBlockFormat(fmt);
    setTextCursor(cursor);
}

void DocumentEditor::indentLess()
{
    QTextCursor cursor = textCursor();
    int indent = cursor.blockFormat().indent();
    if (indent > 0) {
        QTextBlockFormat fmt;
        fmt.setIndent(indent - 1);
        cursor.mergeBlockFormat(fmt);
        setTextCursor(cursor);
    }
}

// ============================================================
// Line Spacing
// ============================================================

void DocumentEditor::setLineSpacing(qreal spacing)
{
    QTextBlockFormat fmt;
    fmt.setLineHeight(static_cast<int>(spacing * 100),
        QTextBlockFormat::ProportionalHeight);
    QTextCursor cursor = textCursor();
    cursor.mergeBlockFormat(fmt);
    setTextCursor(cursor);
}

// ============================================================
// Page Orientation
// ============================================================

void DocumentEditor::togglePageOrientation()
{
    QSizeF size = document()->pageSize();
    qreal temp = size.width();
    size.setWidth(size.height());
    size.setHeight(temp);
    document()->setPageSize(size);
}

// ============================================================
// Page Background
// ============================================================

QColor DocumentEditor::pageBackground() const
{
    return m_pageBackground;
}

void DocumentEditor::setPageBackground(const QColor& color)
{
    m_pageBackground = color;
    QPalette p = viewport()->palette();
    p.setColor(QPalette::Base, color);
    viewport()->setPalette(p);
}

// ============================================================
// Text Direction
// ============================================================

void DocumentEditor::setTextDirection(Qt::LayoutDirection direction)
{
    QTextBlockFormat fmt;
    fmt.setLayoutDirection(direction);
    QTextCursor cursor = textCursor();
    cursor.mergeBlockFormat(fmt);
    setTextCursor(cursor);
}

// ============================================================
// Case
// ============================================================

void DocumentEditor::toUpperCase()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    if (!cursor.hasSelection())
        cursor.select(QTextCursor::WordUnderCursor);
    cursor.insertText(TextUtils::toUpperCase(cursor.selectedText()));
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toLowerCase()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    if (!cursor.hasSelection())
        cursor.select(QTextCursor::WordUnderCursor);
    cursor.insertText(TextUtils::toLowerCase(cursor.selectedText()));
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toTitleCase()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    if (!cursor.hasSelection())
        cursor.select(QTextCursor::WordUnderCursor);
    cursor.insertText(TextUtils::titleCase(cursor.selectedText()));
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toSentenceCase()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    if (!cursor.hasSelection())
        cursor.select(QTextCursor::WordUnderCursor);
    cursor.insertText(TextUtils::sentenceCase(cursor.selectedText()));
    cursor.endEditBlock();
    setTextCursor(cursor);
}

void DocumentEditor::toToggleCase()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    if (!cursor.hasSelection())
        cursor.select(QTextCursor::WordUnderCursor);
    cursor.insertText(TextUtils::toggleCase(cursor.selectedText()));
    cursor.endEditBlock();
    setTextCursor(cursor);
}

// ============================================================
// Spell Check
// ============================================================

void DocumentEditor::setSpellChecker(std::shared_ptr<SpellChecker> checker)
{
    if (m_spellChecker == checker)
        return;
    m_spellChecker = std::move(checker);
    if (m_spellCheckEnabled) {
        if (m_spellHighlighter) {
            m_spellHighlighter->setDocument(nullptr);
            delete m_spellHighlighter;
        }
        m_spellHighlighter = m_spellChecker
            ? new SpellCheckHighlighter(document(), m_spellChecker)
            : nullptr;
    }
}

SpellChecker* DocumentEditor::spellChecker() const
{
    return m_spellChecker.get();
}

void DocumentEditor::setSpellCheckEnabled(bool enabled)
{
    if (m_spellCheckEnabled == enabled)
        return;
    m_spellCheckEnabled = enabled;
    if (enabled && m_spellChecker && !m_spellHighlighter) {
        m_spellHighlighter = new SpellCheckHighlighter(document(), m_spellChecker);
    } else if (!enabled && m_spellHighlighter) {
        m_spellHighlighter->setDocument(nullptr);
        delete m_spellHighlighter;
        m_spellHighlighter = nullptr;
    }
}

// ============================================================
// File I/O
// ============================================================

bool DocumentEditor::loadFromFile(const QString& filename, QString* errorDetail)
{
    // Validate input filename
    if (filename.isEmpty()) {
        if (errorDetail)
            *errorDetail = tr("Filename cannot be empty.");
        return false;
    }

    // Check for path traversal attempts
    if (!isSafeFilePath(filename)) {
        if (errorDetail)
            *errorDetail = tr("Invalid file path detected.");
        return false;
    }

    QFileInfo fi(filename);
    QFile file(filename);
    if (!file.open(QIODevice::ReadOnly)) {
        if (errorDetail)
            *errorDetail = tr("The file could not be read: %1").arg(file.errorString());
        return false;
    }

    constexpr qint64 kMaxFileSize = 50 * 1024 * 1024;
    constexpr qint64 kMaxFileSizeMB = 50;
    if (file.size() > kMaxFileSize) {
        file.close();
        if (errorDetail)
            *errorDetail = tr("The file is too large to open (max %1 MB).").arg(kMaxFileSizeMB);
        return false;
    }

    QByteArray data = file.readAll();
    file.close();

    // Load into a temporary document so a failed parse never partially mutates
    // the live editor document.
    QTextDocument tempDoc;
    tempDoc.setPageSize(document()->pageSize());
    tempDoc.setDocumentMargin(document()->documentMargin());

    QString ext = fi.suffix().toLower();

    bool ok = false;
    if (ext == QStringLiteral("xaml") || ext == QStringLiteral("dexml")) {
        QString xml = QString::fromUtf8(data);
        ok = XamlConverter::loadFromXaml(xml, &tempDoc, m_pageMargins, m_pageBackground,
                                         fi.absolutePath());
        if (!ok && errorDetail)
            *errorDetail = tr("The file could not be opened as an XAML document.");
    } else if (ext == QStringLiteral("html") || ext == QStringLiteral("htm")) {
        QString html = QString::fromUtf8(data);
        tempDoc.setHtml(html);
        ok = true;
    } else if (ext == QStringLiteral("docx")) {
        ok = DocxConverter::loadFromDocx(data, &tempDoc, m_pageMargins, m_pageBackground);
        if (!ok && errorDetail)
            *errorDetail = tr("The file could not be opened as a Word (.docx) document.");
    } else if (ext == QStringLiteral("rtf")) {
        ok = RtfConverter::loadFromRtf(data, &tempDoc);
        if (!ok && errorDetail)
            *errorDetail = tr("The file could not be opened as a Rich Text Format (.rtf) document.");
    } else if (ext == QStringLiteral("txt")) {
        tempDoc.setPlainText(QString::fromUtf8(data));
        ok = true;
    } else if (ext.isEmpty()) {
        if (isLikelyBinary(data)) {
            if (errorDetail)
                *errorDetail = tr("The file appears to be a binary format that cannot be opened.");
            return false;
        }
        tempDoc.setPlainText(QString::fromUtf8(data));
        ok = true;
    } else {
        if (isLikelyBinary(data)) {
            if (errorDetail)
                *errorDetail = tr("The file format \"%1\" is not supported.").arg(ext);
            return false;
        }
        tempDoc.setPlainText(QString::fromUtf8(data));
        ok = true;
    }

    if (!ok) {
        if (errorDetail && errorDetail->isEmpty())
            *errorDetail = tr("The file could not be read.");
        return false;
    }

    // Success: apply the fully-loaded temporary document to the live document.
    restoreDocument(&tempDoc);

    setDocumentName(filename);
    m_fileChanged = false;
    emit modifiedChanged(false);
    document()->setModified(false);
    updateStats();

    return true;
}

void DocumentEditor::restoreDocument(const QTextDocument *source)
{
    QTextDocument *doc = document();
    const QSizeF pageSize = source->pageSize();
    const qreal margin = source->documentMargin();

    QTextCursor cursor(doc);
    cursor.select(QTextCursor::Document);
    cursor.removeSelectedText();
    cursor.insertFragment(QTextDocumentFragment(source));

    // Page margins live on the root frame's format; page size and document
    // margin alone don't round-trip them, so copy the root frame too.
    doc->rootFrame()->setFrameFormat(source->rootFrame()->frameFormat());
    doc->setPageSize(pageSize);
    doc->setDocumentMargin(margin);
    doc->setModified(false);
}

bool DocumentEditor::isLikelyBinary(const QByteArray &data) const
{
    if (data.isEmpty())
        return false;

    int total = data.size();
    int nullCount = 0;
    int nonPrintableCount = 0;
    int sampleSize = qMin(total, 65536);

    for (int i = 0; i < sampleSize; ++i) {
        unsigned char c = static_cast<unsigned char>(data[i]);
        if (c == '\0')
            ++nullCount;
        else if (c < 8 && c != '\t' && c != '\r' && c != '\n')
            ++nonPrintableCount;
    }

    if (nullCount * 100 / sampleSize > 5)
        return true;

    if (nonPrintableCount * 100 / sampleSize > 30)
        return true;

    return false;
}

bool DocumentEditor::saveToFile(const QString& filename, bool updateDocumentName, QString* errorDetail)
{
    auto fail = [&errorDetail](const QString& message) {
        if (errorDetail)
            *errorDetail = message;
        return false;
    };

    // Validate input filename
    if (filename.isEmpty()) {
        qWarning() << "Cannot save to empty filename";
        return fail(tr("Cannot save to an empty filename."));
    }

    // Check for path traversal attempts
    if (!isSafeFilePath(filename)) {
        qWarning() << "Invalid file path detected in saveToFile:" << filename;
        return fail(tr("The file path is invalid."));
    }

    QFileInfo fi(filename);
    QString ext = fi.suffix().toLower();

    // Serialize to bytes before touching the destination file, so a converter
    // failure never truncates the user's previous file.
    QByteArray payload;
    bool allowEmptyPayload = false;
    if (ext == QStringLiteral("xaml") || ext == QStringLiteral("dexml")) {
        payload = XamlConverter::saveToXaml(document(), m_pageMargins, m_pageBackground).toUtf8();
    } else if (ext == QStringLiteral("html") || ext == QStringLiteral("htm")) {
        payload = toHtml().toUtf8();
        allowEmptyPayload = true;
    } else if (ext == QStringLiteral("docx")) {
        payload = DocxConverter::saveToDocx(document(), m_pageMargins, m_pageBackground);
    } else if (ext == QStringLiteral("rtf")) {
        payload = RtfConverter::saveToRtf(document());
    } else {
        payload = toPlainText().toUtf8();
        allowEmptyPayload = true;
    }

    if (payload.isEmpty() && !allowEmptyPayload) {
        qWarning() << "Converter produced no output for" << filename;
        return fail(tr("The document could not be converted to the requested format."));
    }

    // QSaveFile writes to a sibling temp file and atomically renames on commit(),
    // so the destination is never left half-written.
    QSaveFile file(filename);
    if (!file.open(QIODevice::WriteOnly)) {
        qWarning() << "Failed to open file for writing:" << filename << "Error:" << file.errorString();
        return fail(tr("Could not open the file for writing:\n%1").arg(file.errorString()));
    }

    if (file.write(payload) != payload.size() || !file.commit()) {
        qWarning() << "Failed to write file:" << filename << "Error:" << file.errorString();
        return fail(tr("Could not write the file:\n%1").arg(file.errorString()));
    }

    if (updateDocumentName) {
        m_documentName = filename;
        m_fileChanged = false;
        emit modifiedChanged(false);
        document()->setModified(false);
        emit saved(filename);
    }

    return true;
}

// ============================================================
// Events
// ============================================================

void DocumentEditor::keyPressEvent(QKeyEvent* event)
{
    if (event->key() == Qt::Key_Tab) {
        QTextCursor cursor = textCursor();
        QTextBlock block = cursor.block();

        // Inside a list — increase/decrease list indent
        if (QTextList* list = block.textList()) {
            if (event->modifiers() & Qt::ShiftModifier) {
                indentLess();
            } else {
                QTextListFormat fmt = list->format();
                fmt.setIndent(fmt.indent() + 1);
                list->setFormat(fmt);
            }
            return;
        }

        // Inside a table — move to next/previous cell
        QTextTable* table = cursor.currentTable();
        if (table) {
            if (event->modifiers() & Qt::ShiftModifier) {
                QTextCursor prevCursor = cursor;
                prevCursor.movePosition(QTextCursor::PreviousCell);
                if (prevCursor.currentTable() == table)
                    setTextCursor(prevCursor);
                return;
            }
            // If at end of last cell, insert a new row
            int curRow = table->cellAt(cursor).row();
            int curCol = table->cellAt(cursor).column();
            if (curRow == table->rows() - 1 && curCol == table->columns() - 1) {
                table->appendRows(1);
                cursor = table->cellAt(curRow + 1, 0).firstCursorPosition();
                setTextCursor(cursor);
                return;
            }
            QTextCursor nextCursor = cursor;
            nextCursor.movePosition(QTextCursor::NextCell);
            if (nextCursor.currentTable() == table)
                setTextCursor(nextCursor);
            return;
        }

        // Default: increase block indent
        if (event->modifiers() & Qt::ShiftModifier) {
            indentLess();
        } else {
            indentMore();
        }
        return;
    }
    QTextEdit::keyPressEvent(event);
}

