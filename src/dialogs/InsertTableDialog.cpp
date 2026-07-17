#include "InsertTableDialog.h"
#include "widgets/TableGridPicker.h"

#include <QVBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QHBoxLayout>

InsertTableDialog::InsertTableDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Table"));
    setFixedSize(320, 300);

    auto *layout = new QVBoxLayout(this);

    layout->addWidget(new QLabel(tr("Select table size:")));

    m_gridPicker = new TableGridPicker(this);
    layout->addWidget(m_gridPicker, 0, Qt::AlignCenter);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Insert"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(m_gridPicker, &TableGridPicker::tableSelected, this, [this](int r, int c) {
        m_rows = r;
        m_cols = c;
        accept();
    });

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

int InsertTableDialog::rows() const
{
    return m_rows;
}

int InsertTableDialog::columns() const
{
    return m_cols;
}
