#include "SpellCheckHighlighter.h"
#include "utils/SpellChecker.h"

#include <QRegularExpression>

SpellCheckHighlighter::SpellCheckHighlighter(QTextDocument *parent, const SpellChecker *checker)
    : QSyntaxHighlighter(parent)
    , m_checker(checker)
{
}

void SpellCheckHighlighter::highlightBlock(const QString &text)
{
    if (!m_checker || !m_checker->hasDictionary())
        return;

    QRegularExpression wordRegex(QStringLiteral("\\b(\\w+)\\b"));
    QRegularExpressionMatchIterator it = wordRegex.globalMatch(text);

    QTextCharFormat misspelledFormat;
    misspelledFormat.setForeground(Qt::red);
    misspelledFormat.setUnderlineStyle(QTextCharFormat::WaveUnderline);

    while (it.hasNext()) {
        QRegularExpressionMatch match = it.next();
        QString word = match.captured(1);
        if (word.length() < 2)
            continue;
        if (m_checker->isMisspelled(word)) {
            setFormat(match.capturedStart(1), match.capturedLength(1), misspelledFormat);
        }
    }
}
