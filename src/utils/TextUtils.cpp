#include "TextUtils.h"

#include <QRegularExpression>

namespace TextUtils {

QString toUpperCase(const QString &text)
{
    return text.toUpper();
}

QString toLowerCase(const QString &text)
{
    return text.toLower();
}

QString titleCase(const QString &text)
{
    QStringList words = text.toLower().split(QLatin1Char(' '));
    for (QString &word : words) {
        if (!word.isEmpty())
            word[0] = word[0].toUpper();
    }
    return words.join(QLatin1Char(' '));
}

QString sentenceCase(const QString &text)
{
    QString result = text.toLower();
    if (!result.isEmpty())
        result[0] = result[0].toUpper();
    for (int i = 1; i < result.size() - 1; ++i) {
        if (result[i] == QLatin1Char('.') || result[i] == QLatin1Char('!') || result[i] == QLatin1Char('?')) {
            int j = i + 1;
            while (j < result.size() && result[j] == QLatin1Char(' '))
                ++j;
            if (j < result.size())
                result[j] = result[j].toUpper();
        }
    }
    return result;
}

QString toggleCase(const QString &text)
{
    QString result;
    result.reserve(text.size());
    for (const QChar &ch : text) {
        if (ch.isUpper())
            result.append(ch.toLower());
        else if (ch.isLower())
            result.append(ch.toUpper());
        else
            result.append(ch);
    }
    return result;
}

int wordCount(const QString &text)
{
    if (text.trimmed().isEmpty())
        return 0;
    return text.trimmed().split(QRegularExpression(QStringLiteral("\\s+")),
                                Qt::SkipEmptyParts).size();
}

int characterCount(const QString &text, bool includeSpaces)
{
    if (includeSpaces)
        return text.size();
    int count = 0;
    for (const QChar &ch : text) {
        if (!ch.isSpace())
            ++count;
    }
    return count;
}

} // namespace TextUtils
