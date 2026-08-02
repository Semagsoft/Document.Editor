#include "SpellChecker.h"

#include <QFile>
#include <QFileInfo>
#include <QTextStream>
#include <QStandardPaths>
#include <QDir>
#include <QRegularExpression>
#include <QDebug>
#include <algorithm>
#include <climits>

SpellChecker::SpellChecker()
{
    loadSystemDictionary({
        QStringLiteral("/usr/share/dict/words"),
        QStringLiteral("/usr/dict/words"),
        QStringLiteral("/usr/share/dict/american-english"),
        QStringLiteral("/usr/share/dict/british-english")
    });
    loadUserDictionary();
}

SpellChecker::SpellChecker(const QStringList &dictionaryPaths,
                           const QString &userDictionaryPath)
    : m_userDictionaryPath(userDictionaryPath)
{
    loadSystemDictionary(dictionaryPaths);
    loadUserDictionary();
}

void SpellChecker::loadSystemDictionary(const QStringList &candidates)
{
    for (const QString &path : candidates) {
        QFile file(path);
        if (file.open(QIODevice::ReadOnly | QIODevice::Text)) {
            QTextStream stream(&file);
            while (!stream.atEnd()) {
                QString word = stream.readLine().trimmed().toLower();
                if (!word.isEmpty() && word[0].isLetter())
                    m_dictionary.insert(word);
            }
            file.close();
            qDebug() << "SpellChecker: loaded" << m_dictionary.size()
                     << "words from" << path;
            return;
        }
    }

    qDebug() << "SpellChecker: no system dictionary found, spell check disabled";
}

void SpellChecker::loadUserDictionary()
{
    QString path = m_userDictionaryPath;
    if (path.isEmpty())
        path = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation)
               + QStringLiteral("/user_dictionary.txt");
    QFile file(path);
    if (file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        QTextStream stream(&file);
        while (!stream.atEnd()) {
            QString word = stream.readLine().trimmed().toLower();
            if (!word.isEmpty())
                m_userDictionary.insert(word);
        }
        file.close();
    }
}

void SpellChecker::saveUserDictionary()
{
    QString path = m_userDictionaryPath;
    if (path.isEmpty())
        path = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation)
               + QStringLiteral("/user_dictionary.txt");
    QDir().mkpath(QFileInfo(path).absolutePath());
    QFile file(path);
    if (file.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        QTextStream stream(&file);
        for (const QString &word : m_userDictionary)
            stream << word << QStringLiteral("\n");
        file.close();
    }
}

QString SpellChecker::cleanWord(const QString &word) const
{
    QString cleaned;
    for (const QChar &ch : word) {
        if (ch.isLetter() || ch == QLatin1Char('\''))
            cleaned.append(ch.toLower());
    }
    return cleaned;
}

bool SpellChecker::isMisspelled(const QString &word) const
{
    if (!hasDictionary())
        return false;

    QString cleaned = cleanWord(word);
    if (cleaned.isEmpty() || cleaned.length() <= 1)
        return false;

    return !m_dictionary.contains(cleaned) && !m_userDictionary.contains(cleaned);
}

QStringList SpellChecker::findMisspelled(const QString &text) const
{
    QStringList words = text.split(QRegularExpression(QStringLiteral("\\s+")),
                                    Qt::SkipEmptyParts);
    QStringList misspelled;
    for (const QString &word : words) {
        if (isMisspelled(word))
            misspelled.append(word);
    }
    return misspelled;
}

static int levenshteinDistance(const QString &a, const QString &b)
{
    int m = a.length();
    int n = b.length();
    if (m == 0) return n;
    if (n == 0) return m;

    QVector<int> prev(n + 1), curr(n + 1);
    for (int j = 0; j <= n; ++j)
        prev[j] = j;

    for (int i = 1; i <= m; ++i) {
        curr[0] = i;
        for (int j = 1; j <= n; ++j) {
            int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
            curr[j] = std::min({ prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost });
        }
        std::swap(prev, curr);
    }
    return prev[n];
}

static bool byLevenshteinDistance(const QString &word, const QString &a, const QString &b)
{
    return levenshteinDistance(word, a) < levenshteinDistance(word, b);
}

QStringList SpellChecker::suggestions(const QString &word) const
{
    if (!hasDictionary())
        return {};

    QString cleaned = cleanWord(word);
    if (cleaned.isEmpty())
        return {};

    struct Suggestion {
        QString word;
        int distance;
    };

    QVector<Suggestion> scored;
    int len = cleaned.length();
    int maxDistance = qMax(3, len / 2);
    int minLen = qMax(1, len - maxDistance);
    int maxLen = len + maxDistance;
    QChar firstChar = cleaned[0];

    for (const QString &dictWord : m_dictionary) {
        int dlen = dictWord.length();
        if (dlen < minLen || dlen > maxLen)
            continue;
        if (dictWord[0] != firstChar)
            continue;
        int dist = levenshteinDistance(cleaned, dictWord);
        if (dist <= maxDistance)
            scored.append({dictWord, dist});
    }

    std::sort(scored.begin(), scored.end(),
              [](const Suggestion &a, const Suggestion &b) {
                  if (a.distance != b.distance)
                      return a.distance < b.distance;
                  return a.word < b.word;
              });

    QStringList results;
    for (int i = 0; i < scored.size() && i < 10; ++i)
        results.append(scored[i].word);

    return results;
}

void SpellChecker::addToUserDictionary(const QString &word)
{
    QString cleaned = cleanWord(word);
    if (!cleaned.isEmpty()) {
        m_userDictionary.insert(cleaned);
        saveUserDictionary();
    }
}

QStringList SpellChecker::userDictionary() const
{
    return QStringList(m_userDictionary.begin(), m_userDictionary.end());
}

bool SpellChecker::hasDictionary() const
{
    return !m_dictionary.isEmpty();
}
