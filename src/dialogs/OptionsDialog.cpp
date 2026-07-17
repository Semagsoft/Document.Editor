#include "OptionsDialog.h"
#include "app/Settings.h"

#include <QTabWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QCheckBox>
#include <QComboBox>
#include <QLabel>
#include <QSpinBox>
#include <QPushButton>
#include <QTextToSpeech>
#include <QFontComboBox>
#include <QFormLayout>
#include <QGroupBox>

OptionsDialog::OptionsDialog(Settings *settings, QWidget *parent)
    : QDialog(parent)
    , m_settings(settings)
{
    setWindowTitle(tr("Options"));
    setFixedSize(500, 400);

    auto *layout = new QVBoxLayout(this);

    m_tabs = new QTabWidget();
    m_tabs->addTab(createGeneralTab(), tr("General"));
    m_tabs->addTab(createAppearanceTab(), tr("Appearance"));
    m_tabs->addTab(createEditingTab(), tr("Editing"));
    m_tabs->addTab(createTabsTab(), tr("Tabs"));
    m_tabs->addTab(createRulerTab(), tr("Ruler"));
    m_tabs->addTab(createTtsTab(), tr("Text-to-Speech"));
    layout->addWidget(m_tabs);

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("OK"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

void OptionsDialog::accept()
{
    QWidget *ttsTab = m_tabs->widget(5);
    if (ttsTab) {
        auto *voiceCombo = ttsTab->findChild<QComboBox *>();
        auto *speedSpin = ttsTab->findChild<QSpinBox *>();
        if (voiceCombo) m_settings->setTtsVoice(voiceCombo->currentIndex());
        if (speedSpin) m_settings->setTtsSpeed(speedSpin->value());
    }
    QDialog::accept();
}

QWidget *OptionsDialog::createGeneralTab()
{
    auto *widget = new QWidget();
    auto *layout = new QVBoxLayout(widget);

    auto *startupGroup = new QGroupBox(tr("Startup"));
    auto *startupLayout = new QVBoxLayout(startupGroup);
    auto *showStartupCheck = new QCheckBox(tr("Show startup dialog"));
    showStartupCheck->setChecked(m_settings->showStartupDialog());
    startupLayout->addWidget(showStartupCheck);

    auto *updateCheck = new QCheckBox(tr("Check for updates on startup"));
    updateCheck->setChecked(m_settings->checkForUpdatesOnStartup());
    startupLayout->addWidget(updateCheck);

    auto *recentCheck = new QCheckBox(tr("Show recent documents"));
    recentCheck->setChecked(m_settings->showRecentDocuments());
    startupLayout->addWidget(recentCheck);

    auto *startupModeLabel = new QLabel(tr("Startup mode:"));
    auto *startupModeCombo = new QComboBox();
    startupModeCombo->addItems({tr("New document"), tr("Open dialog"), tr("Blank")});
    startupModeCombo->setCurrentIndex(m_settings->startupMode());
    startupLayout->addWidget(startupModeLabel);
    startupLayout->addWidget(startupModeCombo);

    layout->addWidget(startupGroup);

    auto *templatesGroup = new QGroupBox(tr("Templates"));
    auto *templatesLayout = new QVBoxLayout(templatesGroup);
    auto *templatesFolderLabel = new QLabel(tr("Templates folder:"));
    auto *templatesFolderEdit = new QComboBox();
    templatesFolderEdit->setEditable(true);
    templatesFolderEdit->setCurrentText(m_settings->templatesFolder());
    templatesLayout->addWidget(templatesFolderLabel);
    templatesLayout->addWidget(templatesFolderEdit);
    layout->addWidget(templatesGroup);

    layout->addStretch();
    return widget;
}

QWidget *OptionsDialog::createAppearanceTab()
{
    auto *widget = new QWidget();
    auto *layout = new QFormLayout(widget);

    auto *themeCombo = new QComboBox();
    themeCombo->addItems({tr("Office 2010 Blue"), tr("Office 2010 Silver"),
                          tr("Office 2010 Black"), tr("Office 2013"),
                          tr("Windows 8")});
    themeCombo->setCurrentIndex(m_settings->theme());
    layout->addRow(tr("Theme:"), themeCombo);

    auto *fontCombo = new QFontComboBox();
    fontCombo->setCurrentFont(QFont(m_settings->defaultFont()));
    layout->addRow(tr("Default font:"), fontCombo);

    auto *fontSizeSpin = new QSpinBox();
    fontSizeSpin->setRange(8, 72);
    fontSizeSpin->setValue(m_settings->defaultFontSize());
    layout->addRow(tr("Default font size:"), fontSizeSpin);

    auto *glassCheck = new QCheckBox(tr("Enable glass (Aero)"));
    glassCheck->setChecked(m_settings->enableGlass());
    layout->addRow(QString(), glassCheck);

    return widget;
}

QWidget *OptionsDialog::createEditingTab()
{
    auto *widget = new QWidget();
    auto *layout = new QVBoxLayout(widget);

    auto *spellCheck = new QCheckBox(tr("Enable spell check"));
    spellCheck->setChecked(m_settings->spellCheckEnabled());
    layout->addWidget(spellCheck);

    layout->addStretch();
    return widget;
}

QWidget *OptionsDialog::createTabsTab()
{
    auto *widget = new QWidget();
    auto *layout = new QFormLayout(widget);

    auto *placementCombo = new QComboBox();
    placementCombo->addItems({tr("Top"), tr("Bottom"), tr("Left"), tr("Right")});
    placementCombo->setCurrentIndex(m_settings->tabPlacement());
    layout->addRow(tr("Tab placement:"), placementCombo);

    auto *sizeCombo = new QComboBox();
    sizeCombo->addItems({tr("Fit"), tr("Fill"), tr("Fixed")});
    sizeCombo->setCurrentIndex(m_settings->tabSizeMode());
    layout->addRow(tr("Tab size mode:"), sizeCombo);

    auto *closeCombo = new QComboBox();
    closeCombo->addItems({tr("All tabs"), tr("Active tab"), tr("None")});
    closeCombo->setCurrentIndex(m_settings->tabCloseButtonMode());
    layout->addRow(tr("Close button:"), closeCombo);

    return widget;
}

QWidget *OptionsDialog::createRulerTab()
{
    auto *widget = new QWidget();
    auto *layout = new QFormLayout(widget);

    auto *unitCombo = new QComboBox();
    unitCombo->addItems({tr("Inches"), tr("Centimeters")});
    unitCombo->setCurrentIndex(m_settings->rulerMeasurement());
    layout->addRow(tr("Measurement unit:"), unitCombo);

    return widget;
}

QWidget *OptionsDialog::createTtsTab()
{
    auto *widget = new QWidget();
    auto *layout = new QFormLayout(widget);

    auto *voiceCombo = new QComboBox();
    QTextToSpeech tts;
    const QList<QVoice> voices = tts.availableVoices();
    voiceCombo->addItem(tr("Default"));
    for (const QVoice &voice : voices)
        voiceCombo->addItem(voice.name());
    int voiceIdx = m_settings->ttsVoice();
    if (voiceIdx >= 0 && voiceIdx < voiceCombo->count())
        voiceCombo->setCurrentIndex(voiceIdx);
    layout->addRow(tr("Voice:"), voiceCombo);

    auto *speedSpin = new QSpinBox();
    speedSpin->setRange(-10, 10);
    speedSpin->setValue(m_settings->ttsSpeed());
    layout->addRow(tr("Speed:"), speedSpin);

    return widget;
}
