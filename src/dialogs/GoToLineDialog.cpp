#include "GoToLineDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QSpinBox>
#include <QPushButton>

GoToLineDialog::GoToLineDialog(int currentLine, int totalLines, QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Go To Line"));
    setFixedSize(300, 120);

    auto *layout = new QVBoxLayout(this);

    auto *infoLabel = new QLabel(
        tr("Line %1 of %2").arg(currentLine).arg(totalLines));
    layout->addWidget(infoLabel);

    auto *inputLayout = new QHBoxLayout();
    inputLayout->addWidget(new QLabel(tr("Line number:")));

    m_lineSpin = new QSpinBox();
    m_lineSpin->setRange(1, totalLines);
    m_lineSpin->setValue(currentLine);
    inputLayout->addWidget(m_lineSpin);
    layout->addLayout(inputLayout);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Go To"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

int GoToLineDialog::lineNumber() const
{
    return m_lineSpin->value();
}
