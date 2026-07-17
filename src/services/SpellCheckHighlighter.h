#pragma once

#include <QSyntaxHighlighter>
#include <QTextCharFormat>

class SpellChecker;

class SpellCheckHighlighter : public QSyntaxHighlighter
{
    Q_OBJECT

public:
    SpellCheckHighlighter(QTextDocument *parent, SpellChecker *checker);

protected:
    void highlightBlock(const QString &text) override;

private:
    SpellChecker *m_checker;
};
