#include "InsertTimeDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QListWidget>
#include <QPushButton>
#include <QTime>

InsertTimeDialog::InsertTimeDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Time"));
    setFixedSize(300, 220);

    auto *layout = new QVBoxLayout(this);

    layout->addWidget(new QLabel(tr("Select time format:")));

    m_formatList = new QListWidget();
    QStringList formats = {
        QStringLiteral("HH:mm"),
        QStringLiteral("hh:mm AP"),
        QStringLiteral("HH:mm:ss"),
        QStringLiteral("h:mm AP"),
        QStringLiteral("HH:mm:ss.zzz"),
    };
    QTime now = QTime::currentTime();
    for (const auto &fmt : formats)
        m_formatList->addItem(now.toString(fmt));
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

QString InsertTimeDialog::timeFormat() const
{
    static constexpr const char *formats[] = {
        "HH:mm", "hh:mm AP", "HH:mm:ss", "h:mm AP", "HH:mm:ss.zzz"
    };
    int row = m_formatList->currentRow();
    if (row >= 0 && row < 5)
        return QString::fromLatin1(formats[row]);
    return QStringLiteral("HH:mm");
}
