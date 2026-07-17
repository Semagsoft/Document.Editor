#include "FindDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QCheckBox>
#include <QPushButton>

FindDialog::FindDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Find"));
    setFixedSize(380, 200);

    auto *layout = new QVBoxLayout(this);

    auto *searchLayout = new QHBoxLayout();
    searchLayout->addWidget(new QLabel(tr("Find what:")));
    m_searchEdit = new QLineEdit();
    searchLayout->addWidget(m_searchEdit);
    layout->addLayout(searchLayout);

    auto *checkLayout = new QHBoxLayout();
    m_matchCaseCheck = new QCheckBox(tr("Match case"));
    m_wholeWordCheck = new QCheckBox(tr("Whole word"));
    checkLayout->addWidget(m_matchCaseCheck);
    checkLayout->addWidget(m_wholeWordCheck);
    checkLayout->addStretch();
    layout->addLayout(checkLayout);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *findNextBtn = new QPushButton(tr("Find Next"));
    auto *findPrevBtn = new QPushButton(tr("Find Previous"));
    auto *closeBtn = new QPushButton(tr("Close"));
    btnLayout->addWidget(findNextBtn);
    btnLayout->addWidget(findPrevBtn);
    btnLayout->addStretch();
    btnLayout->addWidget(closeBtn);
    layout->addLayout(btnLayout);

    connect(findNextBtn, &QPushButton::clicked, this, [this]() {
        emit findNext(m_searchEdit->text(), m_matchCaseCheck->isChecked(), m_wholeWordCheck->isChecked());
    });
    connect(findPrevBtn, &QPushButton::clicked, this, [this]() {
        emit findPrevious(m_searchEdit->text(), m_matchCaseCheck->isChecked(), m_wholeWordCheck->isChecked());
    });
    connect(closeBtn, &QPushButton::clicked, this, &QDialog::close);
    connect(m_searchEdit, &QLineEdit::returnPressed, this, [this]() {
        emit findNext(m_searchEdit->text(), m_matchCaseCheck->isChecked(), m_wholeWordCheck->isChecked());
    });
}

QString FindDialog::searchText() const
{
    return m_searchEdit->text();
}

bool FindDialog::matchCase() const
{
    return m_matchCaseCheck->isChecked();
}

bool FindDialog::wholeWord() const
{
    return m_wholeWordCheck->isChecked();
}
