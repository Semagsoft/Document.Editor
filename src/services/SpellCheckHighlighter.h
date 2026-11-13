#pragma once

#include <QSyntaxHighlighter>
#include <QTextCharFormat>

#include <memory>

class SpellChecker;

class SpellCheckHighlighter : public QSyntaxHighlighter
{
    Q_OBJECT

public:
    SpellCheckHighlighter(QTextDocument *parent, std::shared_ptr<const SpellChecker> checker);

protected:
    void highlightBlock(const QString &text) override;

private:
    // The checker is shared with the rest of the app; observe it weakly so a
    // dangling pointer can never be dereferenced during repaint.
    std::weak_ptr<const SpellChecker> m_checker;
};
