#include "InsertSymbolDialog.h"
#include "widgets/SymbolPicker.h"

#include <QVBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QHBoxLayout>

InsertSymbolDialog::InsertSymbolDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Symbol"));
    setFixedSize(360, 320);

    auto *layout = new QVBoxLayout(this);

    layout->addWidget(new QLabel(tr("Select a symbol:")));

    m_symbolPicker = new SymbolPicker(this);
    layout->addWidget(m_symbolPicker, 0, Qt::AlignCenter);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Insert"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(m_symbolPicker, &SymbolPicker::symbolSelected, this, [this](const QString &symbol) {
        if (!symbol.isEmpty())
            m_symbol = symbol.at(0);
    });

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

QChar InsertSymbolDialog::selectedSymbol() const
{
    return m_symbol;
}
