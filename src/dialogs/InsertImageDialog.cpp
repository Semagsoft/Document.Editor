#include "InsertImageDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QComboBox>
#include <QFileDialog>
#include <QPixmap>

InsertImageDialog::InsertImageDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Image"));
    setFixedSize(400, 320);

    auto *layout = new QVBoxLayout(this);

    auto *pathLayout = new QHBoxLayout();
    pathLayout->addWidget(new QLabel(tr("Image file:")));
    m_pathEdit = new QLineEdit();
    m_pathEdit->setReadOnly(true);
    pathLayout->addWidget(m_pathEdit);
    auto *browseBtn = new QPushButton(tr("..."));
    browseBtn->setFixedWidth(30);
    pathLayout->addWidget(browseBtn);
    layout->addLayout(pathLayout);

    m_previewLabel = new QLabel(tr("(no preview)"));
    m_previewLabel->setAlignment(Qt::AlignCenter);
    m_previewLabel->setMinimumHeight(120);
    m_previewLabel->setStyleSheet(QStringLiteral("border: 1px solid #ccc;"));
    layout->addWidget(m_previewLabel);

    auto *sizeLayout = new QHBoxLayout();
    sizeLayout->addWidget(new QLabel(tr("Width:")));
    m_widthCombo = new QComboBox();
    m_widthCombo->setEditable(true);
    QStringList sizes = {QStringLiteral("100"), QStringLiteral("200"), QStringLiteral("300"),
                         QStringLiteral("400"), QStringLiteral("500")};
    m_widthCombo->addItems(sizes);
    m_widthCombo->setCurrentText(QStringLiteral("200"));
    sizeLayout->addWidget(m_widthCombo);
    sizeLayout->addWidget(new QLabel(tr("Height:")));
    m_heightCombo = new QComboBox();
    m_heightCombo->setEditable(true);
    m_heightCombo->addItems(sizes);
    m_heightCombo->setCurrentText(QStringLiteral("200"));
    sizeLayout->addWidget(m_heightCombo);
    layout->addLayout(sizeLayout);

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Insert"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(browseBtn, &QPushButton::clicked, this, &InsertImageDialog::browseImage);
    connect(okBtn, &QPushButton::clicked, this, [this]() {
        if (!m_pathEdit->text().isEmpty())
            accept();
    });
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

void InsertImageDialog::browseImage()
{
    QString path = QFileDialog::getOpenFileName(this, tr("Select Image"),
        QString(), tr("Images (*.png *.jpg *.jpeg *.bmp *.gif *.svg);;All Files (*)"));
    if (!path.isEmpty()) {
        m_pathEdit->setText(path);
        updatePreview(path);
    }
}

void InsertImageDialog::updatePreview(const QString &path)
{
    QPixmap pix(path);
    if (!pix.isNull()) {
        m_previewLabel->setPixmap(pix.scaled(200, 120, Qt::KeepAspectRatio, Qt::SmoothTransformation));
        m_widthCombo->setCurrentText(QString::number(pix.width()));
        m_heightCombo->setCurrentText(QString::number(pix.height()));
    }
}

QString InsertImageDialog::imagePath() const
{
    return m_pathEdit->text();
}

int InsertImageDialog::width() const
{
    return m_widthCombo->currentText().toInt();
}

int InsertImageDialog::height() const
{
    return m_heightCombo->currentText().toInt();
}
