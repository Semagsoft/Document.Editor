#include "Settings.h"

Settings::Settings(QObject* parent)
    : QObject(parent)
    , m_settings(QSettings::IniFormat, QSettings::UserScope,
          QStringLiteral("Semagsoft"), QStringLiteral("Document.Editor"))
{
    load();
}

void Settings::load()
{
    m_settings.beginGroup(QStringLiteral("MainWindow"));
    m_windowGeometry = m_settings.value(QStringLiteral("Geometry"),
                                     QRect(100, 100, 1024, 720))
                           .toRect();
    m_windowMaximized = m_settings.value(QStringLiteral("Maximized"), false).toBool();
    m_showRuler = m_settings.value(QStringLiteral("ShowRuler"), true).toBool();
    m_mainWindowState = m_settings.value(QStringLiteral("MainWindowState")).toByteArray();
    m_settings.endGroup();

    m_settings.beginGroup(QStringLiteral("Options"));
    m_startupMode = m_settings.value(QStringLiteral("StartupMode"), 0).toInt();
    m_showStartupDialog = m_settings.value(QStringLiteral("ShowStartupDialog"), true).toBool();
    m_checkForUpdatesOnStartup = m_settings.value(QStringLiteral("CheckForUpdatesOnStartup"), true).toBool();
    m_theme = m_settings.value(QStringLiteral("Theme"), 0).toInt();
    m_enableGlass = m_settings.value(QStringLiteral("EnableGlass"), false).toBool();

    QString family = m_settings.value(QStringLiteral("DefaultFont"), QString()).toString();
    if (!family.isEmpty())
        m_defaultFont = QFont(family);
    else
        m_defaultFont = QFont(QStringLiteral("Segoe UI"));
    m_defaultFontSize = m_settings.value(QStringLiteral("DefaultFontSize"), 12).toInt();

    m_spellCheckEnabled = m_settings.value(QStringLiteral("SpellCheck"), false).toBool();
    m_tabPlacement = m_settings.value(QStringLiteral("TabPlacement"), 0).toInt();
    m_tabSizeMode = m_settings.value(QStringLiteral("TabSizeMode"), 0).toInt();
    m_tabCloseButtonMode = m_settings.value(QStringLiteral("TabCloseButtonMode"), 0).toInt();
    m_showRecentDocuments = m_settings.value(QStringLiteral("ShowRecentDocuments"), true).toBool();
    m_rulerMeasurement = m_settings.value(QStringLiteral("RulerMeasurement"), 0).toInt();
    m_templatesFolder = m_settings.value(QStringLiteral("TemplatesFolder"), QString()).toString();
    m_ttsVoice = m_settings.value(QStringLiteral("TTSVoice"), 0).toInt();
    m_ttsSpeed = m_settings.value(QStringLiteral("TTSSpeed"), 0).toInt();
    m_pluginsEnabled = m_settings.value(QStringLiteral("EnablePlugins"), false).toBool();

    m_recentFiles = m_settings.value(QStringLiteral("RecentFiles"), QStringList()).toStringList();
    m_settings.endGroup();

    emit settingsChanged();
}

void Settings::save()
{
    m_settings.beginGroup(QStringLiteral("MainWindow"));
    m_settings.setValue(QStringLiteral("Geometry"), m_windowGeometry);
    m_settings.setValue(QStringLiteral("Maximized"), m_windowMaximized);
    m_settings.setValue(QStringLiteral("ShowRuler"), m_showRuler);
    m_settings.setValue(QStringLiteral("MainWindowState"), m_mainWindowState);
    m_settings.endGroup();

    m_settings.beginGroup(QStringLiteral("Options"));
    m_settings.setValue(QStringLiteral("StartupMode"), m_startupMode);
    m_settings.setValue(QStringLiteral("ShowStartupDialog"), m_showStartupDialog);
    m_settings.setValue(QStringLiteral("CheckForUpdatesOnStartup"), m_checkForUpdatesOnStartup);
    m_settings.setValue(QStringLiteral("Theme"), m_theme);
    m_settings.setValue(QStringLiteral("EnableGlass"), m_enableGlass);
    m_settings.setValue(QStringLiteral("DefaultFont"), m_defaultFont.family());
    m_settings.setValue(QStringLiteral("DefaultFontSize"), m_defaultFontSize);
    m_settings.setValue(QStringLiteral("SpellCheck"), m_spellCheckEnabled);
    m_settings.setValue(QStringLiteral("TabPlacement"), m_tabPlacement);
    m_settings.setValue(QStringLiteral("TabSizeMode"), m_tabSizeMode);
    m_settings.setValue(QStringLiteral("TabCloseButtonMode"), m_tabCloseButtonMode);
    m_settings.setValue(QStringLiteral("ShowRecentDocuments"), m_showRecentDocuments);
    m_settings.setValue(QStringLiteral("RulerMeasurement"), m_rulerMeasurement);
    m_settings.setValue(QStringLiteral("TemplatesFolder"), m_templatesFolder);
    m_settings.setValue(QStringLiteral("TTSVoice"), m_ttsVoice);
    m_settings.setValue(QStringLiteral("TTSSpeed"), m_ttsSpeed);
    m_settings.setValue(QStringLiteral("EnablePlugins"), m_pluginsEnabled);
    m_settings.setValue(QStringLiteral("RecentFiles"), m_recentFiles);
    m_settings.endGroup();

    m_settings.sync();
    emit settingsChanged();
}

QRect Settings::windowGeometry() const { return m_windowGeometry; }
void Settings::setWindowGeometry(const QRect& rect)
{
    m_windowGeometry = rect;
    emit settingsChanged();
}
bool Settings::windowMaximized() const { return m_windowMaximized; }
void Settings::setWindowMaximized(bool maximized)
{
    m_windowMaximized = maximized;
    emit settingsChanged();
}
bool Settings::showRuler() const { return m_showRuler; }
void Settings::setShowRuler(bool show)
{
    m_showRuler = show;
    emit settingsChanged();
}
QByteArray Settings::mainWindowState() const { return m_mainWindowState; }
void Settings::setMainWindowState(const QByteArray& state) { m_mainWindowState = state; }

int Settings::startupMode() const { return m_startupMode; }
void Settings::setStartupMode(int mode)
{
    m_startupMode = mode;
    emit settingsChanged();
}
bool Settings::showStartupDialog() const { return m_showStartupDialog; }
void Settings::setShowStartupDialog(bool show)
{
    m_showStartupDialog = show;
    emit settingsChanged();
}
bool Settings::checkForUpdatesOnStartup() const { return m_checkForUpdatesOnStartup; }
void Settings::setCheckForUpdatesOnStartup(bool check)
{
    m_checkForUpdatesOnStartup = check;
    emit settingsChanged();
}

int Settings::theme() const { return m_theme; }
void Settings::setTheme(int theme)
{
    m_theme = theme;
    emit settingsChanged();
}
bool Settings::enableGlass() const { return m_enableGlass; }
void Settings::setEnableGlass(bool enable)
{
    m_enableGlass = enable;
    emit settingsChanged();
}

QFont Settings::defaultFont() const { return m_defaultFont; }
void Settings::setDefaultFont(const QFont& font)
{
    m_defaultFont = font;
    emit settingsChanged();
}
int Settings::defaultFontSize() const { return m_defaultFontSize; }
void Settings::setDefaultFontSize(int size)
{
    m_defaultFontSize = size;
    emit settingsChanged();
}

bool Settings::spellCheckEnabled() const { return m_spellCheckEnabled; }
void Settings::setSpellCheckEnabled(bool enabled)
{
    m_spellCheckEnabled = enabled;
    emit settingsChanged();
}

int Settings::tabPlacement() const { return m_tabPlacement; }
void Settings::setTabPlacement(int placement)
{
    m_tabPlacement = placement;
    emit settingsChanged();
}
int Settings::tabSizeMode() const { return m_tabSizeMode; }
void Settings::setTabSizeMode(int mode)
{
    m_tabSizeMode = mode;
    emit settingsChanged();
}
int Settings::tabCloseButtonMode() const { return m_tabCloseButtonMode; }
void Settings::setTabCloseButtonMode(int mode)
{
    m_tabCloseButtonMode = mode;
    emit settingsChanged();
}

QStringList Settings::recentFiles() const { return m_recentFiles; }
void Settings::setRecentFiles(const QStringList& files)
{
    m_recentFiles = files;
    emit settingsChanged();
}

void Settings::addRecentFile(const QString& file)
{
    m_recentFiles.removeAll(file);
    m_recentFiles.prepend(file);
    while (m_recentFiles.size() > kMaxRecentFiles)
        m_recentFiles.removeLast();
    emit settingsChanged();
}

void Settings::clearRecentFiles()
{
    m_recentFiles.clear();
    emit settingsChanged();
}
bool Settings::showRecentDocuments() const { return m_showRecentDocuments; }
void Settings::setShowRecentDocuments(bool show)
{
    m_showRecentDocuments = show;
    emit settingsChanged();
}

int Settings::rulerMeasurement() const { return m_rulerMeasurement; }
void Settings::setRulerMeasurement(int measurement)
{
    m_rulerMeasurement = measurement;
    emit settingsChanged();
}

QString Settings::templatesFolder() const { return m_templatesFolder; }
void Settings::setTemplatesFolder(const QString& folder)
{
    m_templatesFolder = folder;
    emit settingsChanged();
}

int Settings::ttsVoice() const { return m_ttsVoice; }
void Settings::setTtsVoice(int voice)
{
    m_ttsVoice = voice;
    emit settingsChanged();
}
int Settings::ttsSpeed() const { return m_ttsSpeed; }
void Settings::setTtsSpeed(int speed)
{
    m_ttsSpeed = speed;
    emit settingsChanged();
}

bool Settings::pluginsEnabled() const { return m_pluginsEnabled; }
void Settings::setPluginsEnabled(bool enabled)
{
    m_pluginsEnabled = enabled;
    emit settingsChanged();
}
