#pragma once

#include <QMainWindow>
#include <memory>

class QMdiArea;
class QTimer;

class ActionManager;
class DocumentService;
class Settings;
class ThemeManager;
class StatusBarManager;
class DocumentManager;
class DocumentEditor;
class DocumentTab;
class PluginManager;
class PluginContext;
class SpellChecker;
class QTextToSpeech;

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

    void closeCurrentDocument();
    void closeAllDocuments();
    void closeAllButCurrent();

    Settings* settings() const;
    StatusBarManager* statusBarManager() const;

    ActionManager* actionManager() const;

signals:
    void documentChanged();

protected:
    void closeEvent(QCloseEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;

private:
    void createCentralArea();
    void createStatusBar();

    void connectActions();
    void connectFileActions();
    void connectEditActions();
    void connectFormatActions();
    void connectInsertActions();
    void connectPageLayoutActions();
    void connectReviewActions();
    void connectViewActions();
    void connectHelpActions();
    void connectDocumentSignals();
    void loadSettings();
    void saveSettings();
    void applyTabSettings();
    void updateWindowTitle();
    void rebuildRecentFilesMenu();

    DocumentEditor* currentEditor() const;

    // Core widgets
    QMdiArea* m_mdiArea = nullptr;

    // Managers
    ActionManager* m_actionManager = nullptr;
    DocumentService* m_docService = nullptr;
    ThemeManager* m_themeManager = nullptr;
    StatusBarManager* m_statusBarManager = nullptr;
    DocumentManager* m_docManager = nullptr;
    PluginManager* m_pluginManager = nullptr;
    PluginContext* m_pluginContext = nullptr;
    std::unique_ptr<SpellChecker> m_spellChecker;
    QTextToSpeech* m_tts = nullptr;

    Settings* m_settings = nullptr;

    QTimer* m_zoomDebounceTimer = nullptr;
    qreal m_zoomPendingLevel = 1.0;
};
