#include "InsertLinkDialog.h"

#include <QMessageBox>
#include <QUrl>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>

InsertLinkDialog::InsertLinkDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Link"));
    setFixedSize(400, 160);

    auto *layout = new QVBoxLayout(this);

    auto *urlLayout = new QHBoxLayout();
    urlLayout->addWidget(new QLabel(tr("URL:")));
    m_urlEdit = new QLineEdit();
    m_urlEdit->setPlaceholderText(tr("https://example.com"));
    urlLayout->addWidget(m_urlEdit);
    layout->addLayout(urlLayout);

    auto *textLayout = new QHBoxLayout();
    textLayout->addWidget(new QLabel(tr("Display text:")));
    m_textEdit = new QLineEdit();
    m_textEdit->setPlaceholderText(tr("Link text"));
    textLayout->addWidget(m_textEdit);
    layout->addLayout(textLayout);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Insert"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(okBtn, &QPushButton::clicked, this, [this]() {
        if (!m_urlEdit->text().isEmpty())
            accept();
    });
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

void InsertLinkDialog::accept()
{
    QString input = m_urlEdit->text().trimmed();
    if (input.isEmpty())
        return;

    QUrl parsed(input);
    QString scheme = parsed.scheme().toLower();
    QStringList allowedSchemes = {
        QStringLiteral("http"),
        QStringLiteral("https"),
        QStringLiteral("ftp"),
        QStringLiteral("mailto")
    };

    if (!scheme.isEmpty() && !allowedSchemes.contains(scheme)) {
        QMessageBox::warning(this, tr("Invalid URL"),
            tr("URL scheme '%1' is not allowed. Only http, https, ftp, and mailto are supported.")
                .arg(scheme));
        return;
    }

    QDialog::accept();
}

QString InsertLinkDialog::url() const
{
    return m_urlEdit->text();
}

QString InsertLinkDialog::displayText() const
{
    return m_textEdit->text().isEmpty() ? m_urlEdit->text() : m_textEdit->text();
}
