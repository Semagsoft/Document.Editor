#pragma once

#include <QDialog>

class QTabWidget;
class QCheckBox;
class QComboBox;
class QFontComboBox;
class QSpinBox;
class Settings;

class OptionsDialog : public QDialog
{
    Q_OBJECT

public:
    explicit OptionsDialog(Settings *settings, QWidget *parent = nullptr);

    void accept() override;

private:
    Settings *m_settings = nullptr;
    QTabWidget *m_tabs = nullptr;

    QCheckBox *m_showStartupCheck = nullptr;
    QCheckBox *m_updateCheck = nullptr;
    QCheckBox *m_recentCheck = nullptr;
    QComboBox *m_startupModeCombo = nullptr;
    QComboBox *m_templatesFolderEdit = nullptr;

    QComboBox *m_themeCombo = nullptr;
    QFontComboBox *m_fontCombo = nullptr;
    QSpinBox *m_fontSizeSpin = nullptr;
    QCheckBox *m_glassCheck = nullptr;

    QCheckBox *m_spellCheck = nullptr;

    QComboBox *m_tabPlacementCombo = nullptr;
    QComboBox *m_tabSizeCombo = nullptr;
    QComboBox *m_tabCloseCombo = nullptr;

    QComboBox *m_rulerUnitCombo = nullptr;

    QComboBox *m_voiceCombo = nullptr;
    QSpinBox *m_speedSpin = nullptr;

    QWidget *createGeneralTab();
    QWidget *createAppearanceTab();
    QWidget *createEditingTab();
    QWidget *createTabsTab();
    QWidget *createRulerTab();
    QWidget *createTtsTab();
};
