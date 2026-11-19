#include "InsertVideoDialog.h"

#include <QMessageBox>
#include <QUrl>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QFileDialog>

InsertVideoDialog::InsertVideoDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Video"));
    setFixedSize(420, 140);

    auto *layout = new QVBoxLayout(this);

    auto *pathLayout = new QHBoxLayout();
    pathLayout->addWidget(new QLabel(tr("Video file/URL:")));
    m_pathEdit = new QLineEdit();
    m_pathEdit->setPlaceholderText(tr("Enter URL or browse for file"));
    pathLayout->addWidget(m_pathEdit);
    auto *browseBtn = new QPushButton(tr("..."));
    browseBtn->setFixedWidth(30);
    pathLayout->addWidget(browseBtn);
    layout->addLayout(pathLayout);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Insert"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(browseBtn, &QPushButton::clicked, this, &InsertVideoDialog::browseVideo);
    connect(okBtn, &QPushButton::clicked, this, [this]() {
        if (!m_pathEdit->text().isEmpty())
            accept();
    });
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

void InsertVideoDialog::browseVideo()
{
    QString path = QFileDialog::getOpenFileName(this, tr("Select Video"),
        QString(), tr("Videos (*.mp4 *.avi *.mkv *.mov *.wmv);;All Files (*)"));
    if (!path.isEmpty())
        m_pathEdit->setText(path);
}

void InsertVideoDialog::accept()
{
    QString input = m_pathEdit->text().trimmed();
    if (input.isEmpty())
        return;

    // Schemeless input ("example.com/video.mp4") would be treated as relative;
    // assume https unless it looks like a local path.
    if (QUrl(input).scheme().isEmpty()) {
        const bool looksLocal = input.startsWith(QStringLiteral("/"))
            || input.startsWith(QStringLiteral("./"))
            || input.startsWith(QStringLiteral("../"))
            || (input.size() > 1 && input[1] == QLatin1Char(':')); // C:\...
        if (!looksLocal) {
            input = QStringLiteral("https://") + input;
            m_pathEdit->setText(input);
        }
    }

    QUrl parsed(input);
    QString scheme = parsed.scheme().toLower();
    QStringList allowedSchemes = {
        QStringLiteral("http"),
        QStringLiteral("https"),
        QStringLiteral("ftp"),
        QStringLiteral("file")
    };

    if (!scheme.isEmpty() && !allowedSchemes.contains(scheme)) {
        QMessageBox::warning(this, tr("Invalid URL"),
            tr("URL scheme '%1' is not allowed. Only http, https, ftp, and file are supported.")
                .arg(scheme));
        return;
    }

    QDialog::accept();
}

QString InsertVideoDialog::videoPath() const
{
    return m_pathEdit->text().trimmed();
}
