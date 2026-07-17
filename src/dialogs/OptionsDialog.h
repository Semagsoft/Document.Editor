#pragma once

#include <QDialog>

class QTabWidget;
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

    QWidget *createGeneralTab();
    QWidget *createAppearanceTab();
    QWidget *createEditingTab();
    QWidget *createTabsTab();
    QWidget *createRulerTab();
    QWidget *createTtsTab();
};
