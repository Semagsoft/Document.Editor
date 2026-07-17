#include "AboutDialog.h"

#include <QVBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QApplication>
#include <QFont>

AboutDialog::AboutDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("About Document.Editor"));
    setFixedSize(400, 280);

    auto *layout = new QVBoxLayout(this);
    layout->setSpacing(10);

    auto *titleLabel = new QLabel(QStringLiteral("Document.Editor"));
    QFont titleFont = titleLabel->font();
    titleFont.setPointSize(18);
    titleFont.setBold(true);
    titleLabel->setFont(titleFont);
    titleLabel->setAlignment(Qt::AlignCenter);
    layout->addWidget(titleLabel);

    auto *versionLabel = new QLabel(
        tr("Version %1").arg(QApplication::applicationVersion()));
    versionLabel->setAlignment(Qt::AlignCenter);
    layout->addWidget(versionLabel);

    layout->addSpacing(10);

    auto *descLabel = new QLabel(tr(
        "A rich text document editor.\n"
        "Built with Qt %1 and C++17.")
        .arg(QString::fromLatin1(qVersion())));
    descLabel->setAlignment(Qt::AlignCenter);
    descLabel->setWordWrap(true);
    layout->addWidget(descLabel);

    layout->addSpacing(10);

    auto *copyrightLabel = new QLabel(tr("Copyright Semagsoft %1").arg(QStringLiteral("2012-2026")));
    copyrightLabel->setAlignment(Qt::AlignCenter);
    layout->addWidget(copyrightLabel);

    layout->addStretch();

    auto *closeBtn = new QPushButton(tr("Close"));
    connect(closeBtn, &QPushButton::clicked, this, &QDialog::accept);
    layout->addWidget(closeBtn, 0, Qt::AlignCenter);
}
