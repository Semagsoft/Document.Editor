#pragma once

#include <QDialog>
#include <QList>

class QLabel;
class QListWidget;
class QLineEdit;
class QPushButton;
class SpellChecker;

class SpellCheckDialog : public QDialog
{
    Q_OBJECT

public:
    struct Replacement {
        QString original;
        QString replacement;
    };

    explicit SpellCheckDialog(SpellChecker *checker,
                              const QStringList &misspelled,
                              QWidget *parent = nullptr);

    QList<Replacement> replacements() const;

private:
    QLabel *m_wordLabel = nullptr;
    QListWidget *m_suggestionsList = nullptr;
    QLineEdit *m_replacementEdit = nullptr;
    SpellChecker *m_spellChecker = nullptr;
    QStringList m_misspelled;
    int m_currentIndex = 0;
    QList<Replacement> m_replacements;

    void showCurrentWord();
};
