#pragma once

#include <QHash>
#include <QMainWindow>

class QMdiArea;
class QTabWidget;
class QStatusBar;
class QToolBar;
class QMenu;
class QAction;
class QComboBox;
class QSlider;
class QLabel;
class QSpinBox;
class QTimer;

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

signals:
    void documentChanged();

protected:
    void closeEvent(QCloseEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;

private:
    void createMenuBar();
    void createToolBars();
    void createStatusBar();
    void createCentralArea();

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
    void connectTabInsertSignals(DocumentTab* tab);
    void loadSettings();
    void saveSettings();
    void applyTabSettings();
    void updateWindowTitle();

    // File actions
    QAction* m_newAction = nullptr;
    QAction* m_openAction = nullptr;
    QAction* m_closeAction = nullptr;
    QAction* m_closeAllAction = nullptr;
    QAction* m_closeAllButThisAction = nullptr;
    QAction* m_saveAction = nullptr;
    QAction* m_saveAsAction = nullptr;
    QAction* m_saveCopyAction = nullptr;
    QAction* m_saveAllAction = nullptr;
    QAction* m_revertAction = nullptr;
    QAction* m_printAction = nullptr;
    QAction* m_pageSetupAction = nullptr;
    QAction* m_exitAction = nullptr;

    // Import/Export actions
    QAction* m_importFtpAction = nullptr;
    QAction* m_importArchiveAction = nullptr;
    QAction* m_importImageAction = nullptr;
    QAction* m_exportWordpressAction = nullptr;
    QAction* m_exportEmailAction = nullptr;
    QAction* m_exportFtpAction = nullptr;
    QAction* m_exportXpsAction = nullptr;
    QAction* m_exportArchiveAction = nullptr;
    QAction* m_exportImageAction = nullptr;
    QAction* m_exportSoundAction = nullptr;

    // Edit actions
    QAction* m_undoAction = nullptr;
    QAction* m_redoAction = nullptr;
    QAction* m_cutAction = nullptr;
    QAction* m_copyAction = nullptr;
    QAction* m_pasteAction = nullptr;
    QAction* m_pasteTextAction = nullptr;
    QAction* m_pasteImageAction = nullptr;
    QAction* m_deleteAction = nullptr;
    QAction* m_selectAllAction = nullptr;
    QAction* m_findAction = nullptr;
    QAction* m_replaceAction = nullptr;
    QAction* m_goToAction = nullptr;

    // Font actions
    QAction* m_boldAction = nullptr;
    QAction* m_italicAction = nullptr;
    QAction* m_underlineAction = nullptr;
    QAction* m_strikethroughAction = nullptr;
    QAction* m_subscriptAction = nullptr;
    QAction* m_superscriptAction = nullptr;
    QAction* m_clearFormattingAction = nullptr;
    QComboBox* m_fontCombo = nullptr;
    QComboBox* m_fontSizeCombo = nullptr;

    // Paragraph actions
    QAction* m_alignLeftAction = nullptr;
    QAction* m_alignCenterAction = nullptr;
    QAction* m_alignRightAction = nullptr;
    QAction* m_alignJustifyAction = nullptr;
    QAction* m_bulletListAction = nullptr;
    QAction* m_numberListAction = nullptr;
    QAction* m_indentMoreAction = nullptr;
    QAction* m_indentLessAction = nullptr;
    QAction* m_lineSpacingAction = nullptr;
    QAction* m_leftToRightAction = nullptr;
    QAction* m_rightToLeftAction = nullptr;

    // Insert actions
    QAction* m_insertTableAction = nullptr;
    QAction* m_insertImageAction = nullptr;
    QAction* m_insertShapeAction = nullptr;
    QAction* m_insertChartAction = nullptr;
    QAction* m_insertLinkAction = nullptr;
    QAction* m_insertSymbolAction = nullptr;
    QAction* m_insertHorizontalLineAction = nullptr;
    QAction* m_insertDateAction = nullptr;
    QAction* m_insertTimeAction = nullptr;
    QAction* m_insertVideoAction = nullptr;
    QAction* m_insertHeaderAction = nullptr;
    QAction* m_insertFooterAction = nullptr;

    // Page Layout actions
    QAction* m_pageSizeAction = nullptr;
    QAction* m_pageMarginsAction = nullptr;
    QAction* m_pageOrientationAction = nullptr;
    QAction* m_pageBackgroundAction = nullptr;

    // Review actions
    QAction* m_spellCheckAction = nullptr;
    QAction* m_translateAction = nullptr;
    QAction* m_wordCountAction = nullptr;
    QAction* m_ttsAction = nullptr;

    // Options action
    QAction* m_optionsAction = nullptr;

    // View actions
    QAction* m_zoomInAction = nullptr;
    QAction* m_zoomOutAction = nullptr;
    QSlider* m_zoomSlider = nullptr;
    QAction* m_showRulerAction = nullptr;
    QAction* m_showStatusBarAction = nullptr;

    // Help actions
    QAction* m_aboutAction = nullptr;
    QAction* m_checkUpdatesAction = nullptr;
    QAction* m_onlineHelpAction = nullptr;
    QAction* m_websiteAction = nullptr;
    QAction* m_pluginsAction = nullptr;

    // Theme actions
    QAction* m_themeOffice2010Action = nullptr;
    QAction* m_themeOffice2010SilverAction = nullptr;
    QAction* m_themeOffice2010BlackAction = nullptr;
    QAction* m_themeOffice2013Action = nullptr;
    QAction* m_themeWindows8Action = nullptr;

    // Menus
    QMenu* m_fileMenu = nullptr;
    QMenu* m_editMenu = nullptr;
    QMenu* m_insertMenu = nullptr;
    QMenu* m_formatMenu = nullptr;
    QMenu* m_viewMenu = nullptr;
    QMenu* m_helpMenu = nullptr;
    QMenu* m_recentFilesMenu = nullptr;
    QMenu* m_themeMenu = nullptr;
    QMenu* m_importMenu = nullptr;
    QMenu* m_exportMenu = nullptr;

    // Toolbars
    QToolBar* m_quickAccessToolBar = nullptr;
    QToolBar* m_clipboardToolBar = nullptr;
    QToolBar* m_fontToolBar = nullptr;
    QToolBar* m_paragraphToolBar = nullptr;
    QToolBar* m_editingToolBar = nullptr;
    QToolBar* m_insertToolBar = nullptr;
    QToolBar* m_pageLayoutToolBar = nullptr;
    QToolBar* m_reviewToolBar = nullptr;
    QToolBar* m_zoomToolBar = nullptr;

    // Core widgets
    QMdiArea* m_mdiArea = nullptr;

    // Managers
    DocumentService* m_docService = nullptr;
    ThemeManager* m_themeManager = nullptr;
    StatusBarManager* m_statusBarManager = nullptr;
    DocumentManager* m_docManager = nullptr;
    PluginManager* m_pluginManager = nullptr;
    PluginContext* m_pluginContext = nullptr;
    SpellChecker* m_spellChecker = nullptr;
    QTextToSpeech* m_tts = nullptr;

    Settings* m_settings = nullptr;

    bool m_updatingFont = false;
    QTimer* m_zoomDebounceTimer = nullptr;

    void updateEditorActions();
    void rebuildRecentFilesMenu();
    DocumentEditor* currentEditor() const;
};
