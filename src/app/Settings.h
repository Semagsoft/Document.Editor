#pragma once

#include <QByteArray>
#include <QFont>
#include <QObject>
#include <QRect>
#include <QSettings>
#include <QString>
#include <QStringList>

class Settings : public QObject {
    Q_OBJECT

public:
    explicit Settings(QObject* parent = nullptr);

    void load();
    void save();

    // Window geometry
    QRect windowGeometry() const;
    void setWindowGeometry(const QRect& rect);
    bool windowMaximized() const;
    void setWindowMaximized(bool maximized);
    bool showRuler() const;
    void setShowRuler(bool show);
    bool showStatusBar() const;
    void setShowStatusBar(bool show);
    QByteArray mainWindowState() const;
    void setMainWindowState(const QByteArray& state);

    // General options
    int startupMode() const;
    void setStartupMode(int mode);
    bool showStartupDialog() const;
    void setShowStartupDialog(bool show);
    bool checkForUpdatesOnStartup() const;
    void setCheckForUpdatesOnStartup(bool check);

    // Appearance
    int theme() const;
    void setTheme(int theme);
    bool enableGlass() const;
    void setEnableGlass(bool enable);

    // Font
    QFont defaultFont() const;
    void setDefaultFont(const QFont& font);
    int defaultFontSize() const;
    void setDefaultFontSize(int size);

    // Editing
    bool spellCheckEnabled() const;
    void setSpellCheckEnabled(bool enabled);

    // Tabs
    int tabPlacement() const;
    void setTabPlacement(int placement);
    int tabSizeMode() const;
    void setTabSizeMode(int mode);
    int tabCloseButtonMode() const;
    void setTabCloseButtonMode(int mode);

    // Recent files
    QStringList recentFiles() const;
    void setRecentFiles(const QStringList& files);
    void addRecentFile(const QString& file);
    void clearRecentFiles();
    bool showRecentDocuments() const;
    void setShowRecentDocuments(bool show);

    // Ruler
    int rulerMeasurement() const;
    void setRulerMeasurement(int measurement);

    // Templates
    QString templatesFolder() const;
    void setTemplatesFolder(const QString& folder);

    // Text-to-Speech
    int ttsVoice() const;
    void setTtsVoice(int voice);
    int ttsSpeed() const;
    void setTtsSpeed(int speed);

    // Plugins
    bool pluginsEnabled() const;
    void setPluginsEnabled(bool enabled);

signals:
    void settingsChanged();
    void recentFilesChanged();

private:
    static constexpr int kMaxRecentFiles = 10;

    void persistRecentFiles();

    QSettings m_settings;

    // Window
    QRect m_windowGeometry { 100, 100, 1024, 720 };
    bool m_windowMaximized = false;
    bool m_showRuler = true;
    bool m_showStatusBar = true;
    QByteArray m_mainWindowState;

    // General
    int m_startupMode = 0;
    bool m_showStartupDialog = true;
    bool m_checkForUpdatesOnStartup = true;

    // Appearance
    int m_theme = 0;
    bool m_enableGlass = false;

    // Font
    QFont m_defaultFont;
    int m_defaultFontSize = 12;

    // Editing
    bool m_spellCheckEnabled = false;

    // Tabs
    int m_tabPlacement = 0;
    int m_tabSizeMode = 0;
    int m_tabCloseButtonMode = 0;

    // Recent files
    QStringList m_recentFiles;
    bool m_showRecentDocuments = true;

    // Ruler
    int m_rulerMeasurement = 0; // 0 = Inch, 1 = Cm

    // Templates
    QString m_templatesFolder;

    // Text-to-Speech
    int m_ttsVoice = 0;
    int m_ttsSpeed = 0;

    // Plugins
    bool m_pluginsEnabled = false;
};
