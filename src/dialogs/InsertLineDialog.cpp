#include "InsertLineDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QFrame>

InsertLineDialog::InsertLineDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Horizontal Line"));
    setFixedSize(300, 130);

    auto *layout = new QVBoxLayout(this);

    auto *label = new QLabel(tr("Insert a horizontal line at the cursor position."));
    label->setWordWrap(true);
    layout->addWidget(label);

    auto *preview = new QFrame();
    preview->setFrameShape(QFrame::HLine);
    preview->setFrameShadow(QFrame::Sunken);
    layout->addWidget(preview);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Insert"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}
