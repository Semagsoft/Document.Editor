#include "InsertDateDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QListWidget>
#include <QPushButton>
#include <QDate>

InsertDateDialog::InsertDateDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Date"));
    setFixedSize(340, 260);

    auto *layout = new QVBoxLayout(this);

    layout->addWidget(new QLabel(tr("Select date format:")));

    m_formatList = new QListWidget();
    QStringList formats = {
        QStringLiteral("dd/MM/yyyy"),
        QStringLiteral("MM/dd/yyyy"),
        QStringLiteral("yyyy-MM-dd"),
        QStringLiteral("dd-MMM-yyyy"),
        QStringLiteral("MMMM d, yyyy"),
        QStringLiteral("ddd, MMM d, yyyy"),
        QStringLiteral("dddd, MMMM d, yyyy"),
    };
    QDate today = QDate::currentDate();
    for (const auto &fmt : formats)
        m_formatList->addItem(today.toString(fmt));
    m_formatList->setCurrentRow(0);
    layout->addWidget(m_formatList);

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

QString InsertDateDialog::dateFormat() const
{
    static constexpr const char *formats[] = {
        "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd",
        "dd-MMM-yyyy", "MMMM d, yyyy",
        "ddd, MMM d, yyyy", "dddd, MMMM d, yyyy"
    };
    int row = m_formatList->currentRow();
    if (row >= 0 && row < 7)
        return QString::fromLatin1(formats[row]);
    return QStringLiteral("dd/MM/yyyy");
}
