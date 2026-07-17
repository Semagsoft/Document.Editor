#include "ReplaceDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QCheckBox>
#include <QPushButton>

ReplaceDialog::ReplaceDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Replace"));
    setFixedSize(400, 220);

    auto *layout = new QVBoxLayout(this);

    auto *findLayout = new QHBoxLayout();
    findLayout->addWidget(new QLabel(tr("Find what:")));
    m_findEdit = new QLineEdit();
    findLayout->addWidget(m_findEdit);
    layout->addLayout(findLayout);

    auto *replaceLayout = new QHBoxLayout();
    replaceLayout->addWidget(new QLabel(tr("Replace with:")));
    m_replaceEdit = new QLineEdit();
    replaceLayout->addWidget(m_replaceEdit);
    layout->addLayout(replaceLayout);

    m_matchCaseCheck = new QCheckBox(tr("Match case"));
    layout->addWidget(m_matchCaseCheck);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *findNextBtn = new QPushButton(tr("Find Next"));
    auto *replaceBtn = new QPushButton(tr("Replace"));
    auto *replaceAllBtn = new QPushButton(tr("Replace All"));
    auto *closeBtn = new QPushButton(tr("Close"));
    btnLayout->addWidget(findNextBtn);
    btnLayout->addWidget(replaceBtn);
    btnLayout->addWidget(replaceAllBtn);
    btnLayout->addStretch();
    btnLayout->addWidget(closeBtn);
    layout->addLayout(btnLayout);

    connect(findNextBtn, &QPushButton::clicked, this, [this]() {
        emit findNext(m_findEdit->text(), m_matchCaseCheck->isChecked());
    });
    connect(replaceBtn, &QPushButton::clicked, this, [this]() {
        emit replace(m_findEdit->text(), m_replaceEdit->text(), m_matchCaseCheck->isChecked());
    });
    connect(replaceAllBtn, &QPushButton::clicked, this, [this]() {
        emit replaceAll(m_findEdit->text(), m_replaceEdit->text(), m_matchCaseCheck->isChecked());
    });
    connect(closeBtn, &QPushButton::clicked, this, &QDialog::close);
}

QString ReplaceDialog::findText() const
{
    return m_findEdit->text();
}

QString ReplaceDialog::replaceText() const
{
    return m_replaceEdit->text();
}

bool ReplaceDialog::matchCase() const
{
    return m_matchCaseCheck->isChecked();
}
