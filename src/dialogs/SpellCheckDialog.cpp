#include "SpellCheckDialog.h"
#include "utils/SpellChecker.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QListWidget>
#include <QLineEdit>
#include <QPushButton>

SpellCheckDialog::SpellCheckDialog(SpellChecker *checker,
                                   const QStringList &misspelled, QWidget *parent)
    : QDialog(parent)
    , m_spellChecker(checker)
    , m_misspelled(misspelled)
{
    setWindowTitle(tr("Spell Check"));
    setFixedSize(400, 340);

    auto *layout = new QVBoxLayout(this);

    layout->addWidget(new QLabel(tr("Misspelled word:")));
    m_wordLabel = new QLabel();
    m_wordLabel->setStyleSheet(QStringLiteral("font-weight: bold; font-size: 14px;"));
    layout->addWidget(m_wordLabel);

    layout->addWidget(new QLabel(tr("Suggestions:")));
    m_suggestionsList = new QListWidget();
    layout->addWidget(m_suggestionsList);

    auto *replaceLayout = new QHBoxLayout();
    replaceLayout->addWidget(new QLabel(tr("Replace with:")));
    m_replacementEdit = new QLineEdit();
    replaceLayout->addWidget(m_replacementEdit);
    layout->addLayout(replaceLayout);

    auto *btnLayout = new QHBoxLayout();
    auto *replaceBtn = new QPushButton(tr("Replace"));
    auto *ignoreBtn = new QPushButton(tr("Ignore"));
    auto *addBtn = new QPushButton(tr("Add to Dictionary"));
    auto *closeBtn = new QPushButton(tr("Close"));
    btnLayout->addWidget(replaceBtn);
    btnLayout->addWidget(ignoreBtn);
    btnLayout->addWidget(addBtn);
    btnLayout->addStretch();
    btnLayout->addWidget(closeBtn);
    layout->addLayout(btnLayout);

    showCurrentWord();

    connect(m_suggestionsList, &QListWidget::currentTextChanged,
            m_replacementEdit, &QLineEdit::setText);
    connect(m_suggestionsList, &QListWidget::itemDoubleClicked,
            this, [this]() {
        if (m_suggestionsList->currentItem())
            m_replacementEdit->setText(m_suggestionsList->currentItem()->text());
    });
    connect(replaceBtn, &QPushButton::clicked, this, [this]() {
        if (m_currentIndex < m_misspelled.size()) {
            Replacement r;
            r.original = m_misspelled[m_currentIndex];
            r.replacement = m_replacementEdit->text();
            m_replacements.append(r);
        }
        m_currentIndex++;
        if (m_currentIndex < m_misspelled.size())
            showCurrentWord();
        else
            accept();
    });
    connect(ignoreBtn, &QPushButton::clicked, this, [this]() {
        m_currentIndex++;
        if (m_currentIndex < m_misspelled.size())
            showCurrentWord();
        else
            accept();
    });
    connect(addBtn, &QPushButton::clicked, this, [this]() {
        if (m_currentIndex < m_misspelled.size()) {
            m_spellChecker->addToUserDictionary(m_misspelled[m_currentIndex]);
            m_currentIndex++;
            if (m_currentIndex < m_misspelled.size())
                showCurrentWord();
            else
                accept();
        }
    });
    connect(closeBtn, &QPushButton::clicked, this, &QDialog::reject);
}

void SpellCheckDialog::showCurrentWord()
{
    if (m_currentIndex < m_misspelled.size()) {
        QString word = m_misspelled[m_currentIndex];
        m_wordLabel->setText(word);
        m_replacementEdit->setText(word);
        m_suggestionsList->clear();

        QStringList suggs = m_spellChecker ? m_spellChecker->suggestions(word) : QStringList();
        if (suggs.isEmpty()) {
            m_suggestionsList->addItem(word);
        } else {
            for (const QString &s : suggs)
                m_suggestionsList->addItem(s);
        }
    }
}

QList<SpellCheckDialog::Replacement> SpellCheckDialog::replacements() const
{
    return m_replacements;
}
