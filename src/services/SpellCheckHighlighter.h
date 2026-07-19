#pragma once

#include <QSyntaxHighlighter>
#include <QTextCharFormat>

class SpellChecker;

class SpellCheckHighlighter : public QSyntaxHighlighter
{
    Q_OBJECT

public:
    SpellCheckHighlighter(QTextDocument *parent, const SpellChecker *checker);

protected:
    void highlightBlock(const QString &text) override;

private:
    const SpellChecker *m_checker;
};
