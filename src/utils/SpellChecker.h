#pragma once

#include <QString>
#include <QStringList>
#include <QSet>

class SpellChecker
{
public:
    SpellChecker();

    bool isMisspelled(const QString &word) const;
    QStringList findMisspelled(const QString &text) const;
    QStringList suggestions(const QString &word) const;

    void addToUserDictionary(const QString &word);
    QStringList userDictionary() const;

    bool hasDictionary() const;

private:
    QSet<QString> m_dictionary;
    QSet<QString> m_userDictionary;

    void loadSystemDictionary();
    void loadUserDictionary();
    void saveUserDictionary();
    QString cleanWord(const QString &word) const;
};
