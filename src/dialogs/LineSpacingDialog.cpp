#include "LineSpacingDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QDoubleSpinBox>
#include <QListWidget>
#include <QPushButton>
#include <QListWidgetItem>

LineSpacingDialog::LineSpacingDialog(qreal currentSpacing, QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Line Spacing"));
    setFixedSize(320, 280);

    auto *layout = new QVBoxLayout(this);

    auto *presetLabel = new QLabel(tr("Presets:"));
    layout->addWidget(presetLabel);

    m_presetList = new QListWidget();
    m_presetList->addItem(tr("Single (1.0)"));
    m_presetList->addItem(tr("1.15"));
    m_presetList->addItem(tr("1.5"));
    m_presetList->addItem(tr("Double (2.0)"));
    m_presetList->addItem(tr("2.5"));
    m_presetList->addItem(tr("Triple (3.0)"));
    layout->addWidget(m_presetList);

    auto *customLayout = new QHBoxLayout();
    customLayout->addWidget(new QLabel(tr("Custom:")));
    m_spacingSpin = new QDoubleSpinBox();
    m_spacingSpin->setRange(0.5, 10.0);
    m_spacingSpin->setSingleStep(0.1);
    m_spacingSpin->setDecimals(1);
    m_spacingSpin->setValue(currentSpacing);
    customLayout->addWidget(m_spacingSpin);
    layout->addLayout(customLayout);

    layout->addStretch();

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("OK"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(m_presetList, &QListWidget::currentRowChanged, this, [this](int row) {
        static constexpr qreal presets[] = {1.0, 1.15, 1.5, 2.0, 2.5, 3.0};
        if (row >= 0 && row < 6)
            applyPreset(presets[row]);
    });

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

qreal LineSpacingDialog::lineSpacing() const
{
    return m_spacingSpin->value();
}

void LineSpacingDialog::applyPreset(qreal spacing)
{
    m_spacingSpin->setValue(spacing);
}
