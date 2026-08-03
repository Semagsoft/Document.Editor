#include "DocumentTab.h"
#include "DocumentEditor.h"
#include "widgets/RulerWidget.h"
#include "services/DocumentService.h"

#include <QAction>
#include <QApplication>
#include <QColorDialog>
#include <QDate>
#include <QFileInfo>
#include <QFontDialog>
#include <QInputDialog>
#include <QMenu>
#include <QScrollArea>
#include <QTime>
#include <QVBoxLayout>

DocumentTab::DocumentTab(const QString& title, QWidget* parent)
    : QWidget(parent)
    , m_tabTitle(title.isEmpty() ? tr("Untitled") : title)
{
    QVBoxLayout* layout = new QVBoxLayout(this);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setSpacing(0);

    m_ruler = new RulerWidget(this);
    m_ruler->setVisible(false);
    layout->addWidget(m_ruler);

    m_scrollArea = new QScrollArea(this);
    m_scrollArea->setWidgetResizable(true);
    m_scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_scrollArea->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_scrollArea->setFrameShape(QFrame::NoFrame);

    m_editor = new DocumentEditor(this);
    m_editor->setMinimumHeight(200);
    m_editor->setMinimumWidth(200);

    m_scrollArea->setWidget(m_editor);
    layout->addWidget(m_scrollArea);

    createContextMenu();
    updateTitle();

    connect(m_editor, &DocumentEditor::modifiedChanged, this, [this]() {
        updateTitle();
    });

    connect(m_editor, &DocumentEditor::cursorPositionUpdated, this, [this]() {
        updateTitle();
    });
}

DocumentEditor* DocumentTab::editor() const
{
    return m_editor;
}

QString DocumentTab::tabTitle() const
{
    return m_tabTitle;
}

void DocumentTab::setTabTitle(const QString& title)
{
    m_tabTitle = title;
    updateTitle();
}

QString DocumentTab::documentName() const
{
    return m_documentName;
}

void DocumentTab::setDocumentName(const QString& name)
{
    m_documentName = name;
    m_editor->setDocumentName(name);
    QFileInfo fi(name);
    setTabTitle(fi.fileName());
}

void DocumentTab::setPageSize(qreal width, qreal height)
{
    m_editor->setPageWidth(width);
    m_editor->setPageHeight(height);
}

void DocumentTab::updateTitle()
{
    QString title = m_tabTitle;
    if (m_editor && m_editor->isModified())
        title += QStringLiteral(" *");
    emit titleChanged(title);
}

void DocumentTab::createContextMenu()
{
    m_contextMenu = new QMenu(this);

    // Insert submenu (shows dialogs directly to avoid MainWindow coupling)
    QMenu* insertMenu = m_contextMenu->addMenu(tr("Insert"));
    connect(insertMenu->addAction(tr("Table...")), &QAction::triggered,
        this, [this]() { DocumentService::insertTableInteractive(this, m_editor); });
    insertMenu->addSeparator();
    connect(insertMenu->addAction(tr("Image...")), &QAction::triggered,
        this, [this]() { DocumentService::insertImageInteractive(this, m_editor); });
    connect(insertMenu->addAction(tr("Shape...")), &QAction::triggered,
        this, [this]() { DocumentService::insertShapeInteractive(this, m_editor); });
    connect(insertMenu->addAction(tr("Chart...")), &QAction::triggered,
        this, [this]() { DocumentService::insertChartInteractive(this, m_editor); });
    insertMenu->addSeparator();
    connect(insertMenu->addAction(tr("Link...")), &QAction::triggered,
        this, [this]() { DocumentService::insertLinkInteractive(this, m_editor); });
    connect(insertMenu->addAction(tr("Symbol...")), &QAction::triggered,
        this, [this]() { DocumentService::insertSymbolInteractive(this, m_editor); });
    connect(insertMenu->addAction(tr("Horizontal Line...")), &QAction::triggered,
        this, [this]() { DocumentService::insertHorizontalLineInteractive(m_editor); });
    connect(insertMenu->addAction(tr("Date...")), &QAction::triggered,
        this, [this]() { DocumentService::insertDateInteractive(this, m_editor); });
    connect(insertMenu->addAction(tr("Time...")), &QAction::triggered,
        this, [this]() { DocumentService::insertTimeInteractive(this, m_editor); });
    connect(insertMenu->addAction(tr("Video...")), &QAction::triggered,
        this, [this]() { DocumentService::insertVideoInteractive(this, m_editor); });
    insertMenu->addSeparator();
    connect(insertMenu->addAction(tr("Header...")), &QAction::triggered,
        this, [this]() { DocumentService::insertHeader(m_editor); });
    connect(insertMenu->addAction(tr("Footer...")), &QAction::triggered,
        this, [this]() { DocumentService::insertFooter(m_editor); });

    // Format submenu
    QMenu* formatMenu = m_contextMenu->addMenu(tr("Format"));
    connect(formatMenu->addAction(tr("Clear Formatting")), &QAction::triggered,
        m_editor, &DocumentEditor::clearFormatting);

    QMenu* fontMenu = formatMenu->addMenu(tr("Font"));
    connect(fontMenu->addAction(tr("Font Face...")), &QAction::triggered,
        this, [this]() {
            bool ok;
            QFont font = QFontDialog::getFont(&ok, m_editor->currentFont(), this);
            if (ok) {
                QTextCharFormat fmt;
                fmt.setFont(font);
                m_editor->textCursor().mergeCharFormat(fmt);
            }
        });
    connect(fontMenu->addAction(tr("Font Size...")), &QAction::triggered,
        this, [this]() {
            bool ok;
            int size = QInputDialog::getInt(this, tr("Font Size"), tr("Size:"),
                qRound(m_editor->currentFont().pointSizeF()), 1, 999, 1, &ok);
            if (ok) {
                QTextCharFormat fmt;
                fmt.setFontPointSize(size);
                m_editor->textCursor().mergeCharFormat(fmt);
            }
        });
    connect(fontMenu->addAction(tr("Font Color...")), &QAction::triggered,
        this, [this]() {
            QColor color = QColorDialog::getColor(m_editor->textColor(), this, tr("Font Color"));
            if (color.isValid()) {
                QTextCharFormat fmt;
                fmt.setForeground(color);
                m_editor->textCursor().mergeCharFormat(fmt);
            }
        });
    connect(fontMenu->addAction(tr("Highlight Color...")), &QAction::triggered,
        this, [this]() {
            QColor color = QColorDialog::getColor(m_editor->textBackgroundColor(),
                this, tr("Highlight Color"));
            if (color.isValid()) {
                QTextCharFormat fmt;
                fmt.setBackground(color);
                m_editor->textCursor().mergeCharFormat(fmt);
            }
        });

    QMenu* styleMenu = formatMenu->addMenu(tr("Style"));
    QAction* boldAction = styleMenu->addAction(tr("Bold"));
    boldAction->setCheckable(true);
    QAction* italicAction = styleMenu->addAction(tr("Italic"));
    italicAction->setCheckable(true);
    QAction* underlineAction = styleMenu->addAction(tr("Underline"));
    underlineAction->setCheckable(true);
    QAction* strikeAction = styleMenu->addAction(tr("Strikethrough"));
    strikeAction->setCheckable(true);

    QAction* subAction = formatMenu->addAction(tr("Subscript"));
    QAction* superAction = formatMenu->addAction(tr("Superscript"));
    formatMenu->addSeparator();
    QAction* indentMoreAction = formatMenu->addAction(tr("Indent More"));
    QAction* indentLessAction = formatMenu->addAction(tr("Indent Less"));

    QMenu* listMenu = formatMenu->addMenu(tr("List"));
    QAction* bulletListAction = listMenu->addAction(tr("Bullet List"));
    QAction* numberListAction = listMenu->addAction(tr("Number List"));

    QMenu* alignMenu = formatMenu->addMenu(tr("Align"));
    QAction* alignLeftAction = alignMenu->addAction(tr("Left"));
    QAction* alignCenterAction = alignMenu->addAction(tr("Center"));
    QAction* alignRightAction = alignMenu->addAction(tr("Right"));
    QAction* alignJustifyAction = alignMenu->addAction(tr("Justify"));

    QAction* lineSpacingAction = formatMenu->addAction(tr("Line Spacing..."));
    formatMenu->addSeparator();
    QAction* ltrAction = formatMenu->addAction(tr("Left to Right"));
    QAction* rtlAction = formatMenu->addAction(tr("Right to Left"));

    QMenu* pageMenu = m_contextMenu->addMenu(tr("Page"));
    QAction* pageSizeAction = pageMenu->addAction(tr("Page Size..."));
    QAction* pageMarginsAction = pageMenu->addAction(tr("Page Margins..."));
    QAction* pageOrientationAction = pageMenu->addAction(tr("Orientation"));

    m_contextMenu->addSeparator();

    // Edit actions
    QAction* undoAction = m_contextMenu->addAction(tr("Undo"));
    undoAction->setShortcut(QKeySequence::Undo);
    QAction* redoAction = m_contextMenu->addAction(tr("Redo"));
    redoAction->setShortcut(QKeySequence::Redo);

    m_contextMenu->addSeparator();

    QAction* cutAction = m_contextMenu->addAction(tr("Cut"));
    cutAction->setShortcut(QKeySequence::Cut);
    QAction* copyAction = m_contextMenu->addAction(tr("Copy"));
    copyAction->setShortcut(QKeySequence::Copy);
    QAction* pasteAction = m_contextMenu->addAction(tr("Paste"));
    pasteAction->setShortcut(QKeySequence::Paste);
    QAction* deleteAction = m_contextMenu->addAction(tr("Delete"));
    deleteAction->setShortcut(QKeySequence::Delete);

    m_contextMenu->addSeparator();

    QAction* selectAllAction = m_contextMenu->addAction(tr("Select All"));
    m_contextMenu->addSeparator();
    connect(m_contextMenu->addAction(tr("Find...")), &QAction::triggered,
        this, [this]() {
            bool ok;
            QString text = QInputDialog::getText(this, tr("Find"),
                tr("Find:"), QLineEdit::Normal, QString(), &ok);
            if (ok && !text.isEmpty())
                m_editor->find(text);
        });
    connect(m_contextMenu->addAction(tr("Go To...")), &QAction::triggered,
        this, [this]() {
            bool ok;
            int line = QInputDialog::getInt(this, tr("Go To Line"),
                tr("Line number:"), 1, 1, m_editor->lineCount(), 1, &ok);
            if (ok)
                m_editor->goToLine(line);
        });

    // Connect format actions
    connect(boldAction, &QAction::triggered, m_editor, &DocumentEditor::toggleBold);
    connect(italicAction, &QAction::triggered, m_editor, &DocumentEditor::toggleItalic);
    connect(underlineAction, &QAction::triggered, m_editor, &DocumentEditor::toggleUnderline);
    connect(strikeAction, &QAction::triggered, m_editor, &DocumentEditor::toggleStrikethrough);
    connect(subAction, &QAction::triggered, m_editor, &DocumentEditor::toggleSubscript);
    connect(superAction, &QAction::triggered, m_editor, &DocumentEditor::toggleSuperscript);
    connect(indentMoreAction, &QAction::triggered, m_editor, &DocumentEditor::indentMore);
    connect(indentLessAction, &QAction::triggered, m_editor, &DocumentEditor::indentLess);
    connect(bulletListAction, &QAction::triggered, m_editor, &DocumentEditor::toggleBulletList);
    connect(numberListAction, &QAction::triggered, m_editor, &DocumentEditor::toggleNumberList);
    connect(alignLeftAction, &QAction::triggered, this, [this]() {
        m_editor->setParagraphAlignment(Qt::AlignLeft);
    });
    connect(alignCenterAction, &QAction::triggered, this, [this]() {
        m_editor->setParagraphAlignment(Qt::AlignCenter);
    });
    connect(alignRightAction, &QAction::triggered, this, [this]() {
        m_editor->setParagraphAlignment(Qt::AlignRight);
    });
    connect(alignJustifyAction, &QAction::triggered, this, [this]() {
        m_editor->setParagraphAlignment(Qt::AlignJustify);
    });
    connect(lineSpacingAction, &QAction::triggered, this, [this]() {
        qreal current = m_editor->textCursor().blockFormat().lineHeight() / 100.0;
        QInputDialog dlg;
        bool ok;
        double spacing = QInputDialog::getDouble(this, tr("Line Spacing"),
            tr("Spacing (0.5 - 10.0):"), current, 0.5, 10.0, 1, &ok);
        if (ok)
            m_editor->setLineSpacing(spacing);
    });
    connect(ltrAction, &QAction::triggered, this, [this]() {
        m_editor->setTextDirection(Qt::LeftToRight);
    });
    connect(rtlAction, &QAction::triggered, this, [this]() {
        m_editor->setTextDirection(Qt::RightToLeft);
    });
    connect(pageSizeAction, &QAction::triggered, this, [this]() {
        bool ok;
        QString result = QInputDialog::getText(this, tr("Page Size"),
            tr("Width x Height (points):"), QLineEdit::Normal,
            QStringLiteral("%1 x %2")
                .arg(m_editor->pageWidth())
                .arg(m_editor->pageHeight()),
            &ok);
        if (!ok)
            return;
        QStringList parts = result.split(QStringLiteral("x"), Qt::SkipEmptyParts);
        if (parts.size() != 2)
            return;
        bool wOk = false, hOk = false;
        qreal w = parts[0].trimmed().toDouble(&wOk);
        qreal h = parts[1].trimmed().toDouble(&hOk);
        if (wOk && hOk && w > 0 && h > 0)
            m_editor->setPageWidth(w), m_editor->setPageHeight(h);
    });
    connect(pageMarginsAction, &QAction::triggered, this, [this]() {
        QMarginsF m = m_editor->pageMargins();
        QInputDialog dlg;
        bool ok;
        QString text = QInputDialog::getText(this, tr("Page Margins"),
            tr("Left, Top, Right, Bottom (points):"), QLineEdit::Normal,
            QStringLiteral("%1, %2, %3, %4")
                .arg(m.left()).arg(m.top()).arg(m.right()).arg(m.bottom()),
            &ok);
        if (!ok)
            return;
        const QStringList parts = text.split(QLatin1Char(','), Qt::SkipEmptyParts);
        if (parts.size() != 4)
            return;
        QVector<qreal> vals;
        for (const QString &p : parts) {
            bool vOk = false;
            qreal v = p.trimmed().toDouble(&vOk);
            if (!vOk || v < 0)
                return;
            vals.append(v);
        }
        m_editor->setPageMargins(QMarginsF(vals[0], vals[1], vals[2], vals[3]));
    });
    connect(pageOrientationAction, &QAction::triggered, this, [this]() {
        bool ok;
        int cur = m_editor->pageWidth() >= m_editor->pageHeight() ? 1 : 0;
        int idx = QInputDialog::getInt(this, tr("Orientation"),
            tr("0 = Portrait, 1 = Landscape"), cur, 0, 1, 1, &ok);
        if (!ok)
            return;
        qreal w = m_editor->pageWidth();
        qreal h = m_editor->pageHeight();
        if ((idx == 1 && w < h) || (idx == 0 && w > h)) {
            m_editor->setPageWidth(h);
            m_editor->setPageHeight(w);
        }
    });

    // Connect edit actions
    connect(undoAction, &QAction::triggered, m_editor, &QTextEdit::undo);
    connect(redoAction, &QAction::triggered, m_editor, &QTextEdit::redo);
    connect(cutAction, &QAction::triggered, m_editor, &QTextEdit::cut);
    connect(copyAction, &QAction::triggered, m_editor, &QTextEdit::copy);
    connect(pasteAction, &QAction::triggered, m_editor, &QTextEdit::paste);
    connect(deleteAction, &QAction::triggered, m_editor, [this]() {
        m_editor->textCursor().removeSelectedText();
    });
    connect(selectAllAction, &QAction::triggered, m_editor, &QTextEdit::selectAll);

    m_editor->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_editor, &QWidget::customContextMenuRequested, this, [this](const QPoint& pos) {
        m_contextMenu->exec(m_editor->mapToGlobal(pos));
    });
}

RulerWidget* DocumentTab::ruler() const
{
    return m_ruler;
}

void DocumentTab::setRulerVisible(bool visible)
{
    m_ruler->setVisible(visible);
}

void DocumentTab::setRulerUnit(RulerWidget::Unit unit)
{
    m_ruler->setUnit(unit);
}
