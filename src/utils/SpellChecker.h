#pragma once

#include <QString>
#include <QStringList>
#include <QSet>

class SpellChecker
{
public:
    SpellChecker();

    // For tests: empty dictionaryPaths means "no system dictionary".
    // userDictionaryPath overrides the default user dictionary location.
    explicit SpellChecker(const QStringList &dictionaryPaths,
                          const QString &userDictionaryPath = QString());

    bool isMisspelled(const QString &word) const;
    QStringList findMisspelled(const QString &text) const;
    QStringList suggestions(const QString &word) const;

    void addToUserDictionary(const QString &word);
    QStringList userDictionary() const;

    bool hasDictionary() const;

private:
    QSet<QString> m_dictionary;
    QSet<QString> m_userDictionary;
    QString m_userDictionaryPath;

    void loadSystemDictionary(const QStringList &candidates);
    void loadUserDictionary();
    void saveUserDictionary();
    QString cleanWord(const QString &word) const;
};
