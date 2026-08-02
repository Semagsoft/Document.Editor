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
#include <QFile>
#include <QFileInfo>
#include <QFontDatabase>
#include <QKeyEvent>
#include <QMenu>
#include <QMessageBox>
#include <QMimeData>
#include <QRegularExpression>
#include <QTextBlock>
#include <QTextCursor>
#include <QTextList>
#include <QTextTable>

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
    if (fi.exists())
        m_readOnlyFile = !fi.isWritable();
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
    m_zoomLevel = qBound(0.1, level, 5.0);
    QFont zoomedFont = m_baseFont;
    zoomedFont.setPointSizeF(qRound(m_baseFontPointSize * m_zoomLevel));
    setFont(zoomedFont);
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
        maxCol = block.length();

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
}

void DocumentEditor::toggleItalic()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setFontItalic(!cursor.charFormat().fontItalic());
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
}

void DocumentEditor::toggleUnderline()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setFontUnderline(!cursor.charFormat().fontUnderline());
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
}

void DocumentEditor::toggleStrikethrough()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    QTextCharFormat fmt;
    fmt.setFontStrikeOut(!cursor.charFormat().fontStrikeOut());
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
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
    fmt.setForeground(QColor());
    fmt.setBackground(QColor());
    cursor.mergeCharFormat(fmt);
    cursor.endEditBlock();
}

// ============================================================
// Alignment
// ============================================================

void DocumentEditor::setParagraphAlignment(Qt::Alignment align)
{
    QTextBlockFormat fmt;
    fmt.setAlignment(align);
    textCursor().mergeBlockFormat(fmt);
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
}

// ============================================================
// Indent
// ============================================================

void DocumentEditor::indentMore()
{
    QTextBlockFormat fmt;
    fmt.setIndent(textCursor().blockFormat().indent() + 1);
    textCursor().mergeBlockFormat(fmt);
}

void DocumentEditor::indentLess()
{
    int indent = textCursor().blockFormat().indent();
    if (indent > 0) {
        QTextBlockFormat fmt;
        fmt.setIndent(indent - 1);
        textCursor().mergeBlockFormat(fmt);
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
    textCursor().mergeBlockFormat(fmt);
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
    textCursor().mergeBlockFormat(fmt);
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
}

void DocumentEditor::toLowerCase()
{
    QTextCursor cursor = textCursor();
    cursor.beginEditBlock();
    if (!cursor.hasSelection())
        cursor.select(QTextCursor::WordUnderCursor);
    cursor.insertText(TextUtils::toLowerCase(cursor.selectedText()));
    cursor.endEditBlock();
}

// ============================================================
// Spell Check
// ============================================================

void DocumentEditor::setSpellChecker(SpellChecker* checker)
{
    if (m_spellChecker == checker && checker)
        return;
    m_spellChecker = checker;
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
    return m_spellChecker;
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

bool DocumentEditor::loadFromFile(const QString& filename)
{
    QFile file(filename);
    if (!file.open(QIODevice::ReadOnly))
        return false;

    constexpr qint64 kMaxFileSize = 50 * 1024 * 1024;
    constexpr qint64 kMaxFileSizeMB = 50;
    if (file.size() > kMaxFileSize) {
        file.close();
        QMessageBox::warning(this, tr("File Too Large"),
            tr("The file is too large to open (max %1 MB).").arg(kMaxFileSizeMB));
        return false;
    }

    QByteArray data = file.readAll();
    file.close();

    QFileInfo fi(filename);
    QString ext = fi.suffix().toLower();

    bool ok = false;
    if (ext == QStringLiteral("xaml") || ext == QStringLiteral("dexml")) {
        QString xml = QString::fromUtf8(data);
        ok = XamlConverter::loadFromXaml(xml, document(), m_pageMargins, m_pageBackground);
        if (!ok) {
            setPlainText(xml);
        }
    } else if (ext == QStringLiteral("html") || ext == QStringLiteral("htm")) {
        QString html = QString::fromUtf8(data);
        setHtml(html);
        ok = true;
    } else if (ext == QStringLiteral("docx")) {
        ok = DocxConverter::loadFromDocx(data, document(), m_pageMargins, m_pageBackground);
        if (!ok) {
            QMessageBox::warning(this, tr("Cannot Open File"),
                tr("The file could not be opened as a Word (.docx) document."));
            return false;
        }
    } else if (ext == QStringLiteral("rtf")) {
        ok = RtfConverter::loadFromRtf(data, document());
        if (!ok) {
            QMessageBox::warning(this, tr("Cannot Open File"),
                tr("The file could not be opened as a Rich Text Format (.rtf) document."));
            return false;
        }
    } else if (ext == QStringLiteral("txt")) {
        setPlainText(QString::fromUtf8(data));
        ok = true;
    } else if (ext.isEmpty()) {
        // No extension — treat as plain text but reject binary content
        if (isLikelyBinary(data)) {
            QMessageBox::warning(this, tr("Cannot Open File"),
                tr("The file appears to be a binary format that cannot be opened."));
            return false;
        }
        setPlainText(QString::fromUtf8(data));
        ok = true;
    } else {
        // Unknown extension — treat as plain text but reject binary content
        if (isLikelyBinary(data)) {
            QMessageBox::warning(this, tr("Cannot Open File"),
                tr("The file format \"%1\" is not supported.").arg(ext));
            return false;
        }
        setPlainText(QString::fromUtf8(data));
        ok = true;
    }

    if (ok) {
        m_documentName = filename;
        m_fileChanged = false;
        emit modifiedChanged(false);
        document()->setModified(false);
        updateStats();
    }

    return ok;
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

bool DocumentEditor::saveToFile(const QString& filename)
{
    QFile file(filename);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Truncate))
        return false;

    QFileInfo fi(filename);
    QString ext = fi.suffix().toLower();

    bool ok = false;
    if (ext == QStringLiteral("xaml") || ext == QStringLiteral("dexml")) {
        QByteArray xml = XamlConverter::saveToXaml(document(), m_pageMargins, m_pageBackground).toUtf8();
        ok = file.write(xml) > 0;
    } else if (ext == QStringLiteral("html") || ext == QStringLiteral("htm")) {
        QByteArray html = toHtml().toUtf8();
        ok = file.write(html) > 0;
    } else if (ext == QStringLiteral("docx")) {
        QByteArray docxData = DocxConverter::saveToDocx(document(), m_pageMargins, m_pageBackground);
        ok = file.write(docxData) > 0;
    } else if (ext == QStringLiteral("rtf")) {
        ok = file.write(RtfConverter::saveToRtf(document())) > 0;
    } else {
        QByteArray text = toPlainText().toUtf8();
        ok = file.write(text) > 0;
    }

    file.close();

    if (ok) {
        m_documentName = filename;
        m_fileChanged = false;
        emit modifiedChanged(false);
        document()->setModified(false);
    }

    return ok;
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

void DocumentEditor::contextMenuEvent(QContextMenuEvent* event)
{
    QMenu* menu = createStandardContextMenu();
    menu->addSeparator();

    QAction* upperAction = menu->addAction(tr("Uppercase"));
    connect(upperAction, &QAction::triggered, this, &DocumentEditor::toUpperCase);

    QAction* lowerAction = menu->addAction(tr("Lowercase"));
    connect(lowerAction, &QAction::triggered, this, &DocumentEditor::toLowerCase);

    menu->exec(event->globalPos());
    delete menu;
}
