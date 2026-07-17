#pragma once

#include <QMarginsF>
#include <QString>
#include <QTextBlockFormat>
#include <QTextCharFormat>
#include <QTextCursor>
#include <QTextDocument>
#include <QTextEdit>
#include <QTextListFormat>
#include <QTimer>

class SpellChecker;
class SpellCheckHighlighter;

class DocumentEditor : public QTextEdit {
    Q_OBJECT

public:
    explicit DocumentEditor(QWidget* parent = nullptr);

    // Document state
    QString documentName() const;
    void setDocumentName(const QString& name);
    bool isModified() const;
    void setModified(bool changed);
    bool isReadOnlyFile() const;

    // Page setup
    QMarginsF pageMargins() const;
    void setPageMargins(const QMarginsF& margins);
    qreal pageWidth() const;
    void setPageWidth(qreal width);
    qreal pageHeight() const;
    void setPageHeight(qreal height);

    // Zoom
    qreal zoomLevel() const;
    void setZoomLevel(qreal level);
    void setBaseFontPointSize(qreal size);

    // Statistics
    int lineCount() const;
    int columnCount() const;
    int wordCount() const;
    int selectedLineNumber() const;
    int selectedColumnNumber() const;

    // Find
    QTextCursor findWord(const QString& word, const QTextCursor& start = QTextCursor());
    void goToLine(int lineNumber);

    // Formatting toggles
    void toggleBold();
    void toggleItalic();
    void toggleUnderline();
    void toggleStrikethrough();
    void toggleSubscript();
    void toggleSuperscript();
    void clearFormatting();

    // Alignment
    void setParagraphAlignment(Qt::Alignment align);

    // Lists
    void toggleBulletList();
    void toggleNumberList();

    // Indent
    void indentMore();
    void indentLess();

    // Line spacing
    void setLineSpacing(qreal spacing);

    // Page orientation
    void togglePageOrientation();

    // Page background
    QColor pageBackground() const;
    void setPageBackground(const QColor& color);

    // Text direction
    void setTextDirection(Qt::LayoutDirection direction);

    // Case
    void toUpperCase();
    void toLowerCase();

    // Spell check
    void setSpellChecker(SpellChecker* checker);
    SpellChecker* spellChecker() const;
    void setSpellCheckEnabled(bool enabled);

    // File I/O
    bool loadFromFile(const QString& filename);
    bool saveToFile(const QString& filename);

signals:
    void modifiedChanged(bool changed);
    void cursorPositionUpdated();

protected:
    void keyPressEvent(QKeyEvent* event) override;
    void contextMenuEvent(QContextMenuEvent* event) override;

private:
    void updateStats();
    void connectDocumentSignals();
    void scheduleStatsUpdate();

    QString m_documentName;
    bool m_fileChanged = false;
    bool m_readOnlyFile = false;
    qreal m_zoomLevel = 1.0;
    qreal m_baseFontPointSize = 12.0;
    QFont m_baseFont;
    QMarginsF m_pageMargins { 96, 96, 96, 96 };
    QColor m_pageBackground { Qt::white };
    bool m_pageBackgroundSet = false;

    QTimer* m_statsTimer = nullptr;

    SpellChecker* m_spellChecker = nullptr;
    SpellCheckHighlighter* m_spellHighlighter = nullptr;
    bool m_spellCheckEnabled = false;

    // Cached stats
    int m_lineCount = 0;
    int m_columnCount = 0;
    int m_wordCount = 0;
    int m_selectedLine = 0;
    int m_selectedColumn = 0;
};
