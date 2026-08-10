#pragma once

#include <QObject>

class QAction;
class QComboBox;
class QLabel;
class QMainWindow;
class QMenu;
class QMenuBar;
class QSlider;
class QToolBar;

class DocumentEditor;
class DocumentManager;
class StatusBarManager;

class ActionManager : public QObject {
    Q_OBJECT

public:
    explicit ActionManager(QObject* parent = nullptr);

    void setupActions(QMainWindow* mainWindow);
    void setupToolBars(QMainWindow* mainWindow);
    void updateEditorActions(DocumentEditor* editor, DocumentManager* docManager, StatusBarManager* statusBarManager);

    // File actions
    QAction* newAction() const;
    QAction* openAction() const;
    QAction* closeAction() const;
    QAction* closeAllAction() const;
    QAction* closeAllButThisAction() const;
    QAction* saveAction() const;
    QAction* saveAsAction() const;
    QAction* saveCopyAction() const;
    QAction* saveAllAction() const;
    QAction* revertAction() const;
    QAction* printAction() const;
    QAction* pageSetupAction() const;
    QAction* exitAction() const;

    // Import/Export actions
    QAction* importFtpAction() const;
    QAction* importArchiveAction() const;
    QAction* importImageAction() const;
    QAction* exportWordpressAction() const;
    QAction* exportEmailAction() const;
    QAction* exportFtpAction() const;
    QAction* exportPdfAction() const;
    QAction* exportArchiveAction() const;
    QAction* exportImageAction() const;
    QAction* exportSoundAction() const;

    // Edit actions
    QAction* undoAction() const;
    QAction* redoAction() const;
    QAction* cutAction() const;
    QAction* copyAction() const;
    QAction* pasteAction() const;
    QAction* pasteTextAction() const;
    QAction* pasteImageAction() const;
    QAction* deleteAction() const;
    QAction* selectAllAction() const;
    QAction* findAction() const;
    QAction* replaceAction() const;
    QAction* goToAction() const;

    // Font actions
    QAction* boldAction() const;
    QAction* italicAction() const;
    QAction* underlineAction() const;
    QAction* strikethroughAction() const;
    QAction* subscriptAction() const;
    QAction* superscriptAction() const;
    QAction* clearFormattingAction() const;
    QAction* fontFaceAction() const;
    QAction* fontSizeAction() const;
    QAction* fontColorAction() const;
    QAction* highlightColorAction() const;
    QComboBox* fontCombo() const;
    QComboBox* fontSizeCombo() const;

    // Paragraph actions
    QAction* alignLeftAction() const;
    QAction* alignCenterAction() const;
    QAction* alignRightAction() const;
    QAction* alignJustifyAction() const;
    QAction* bulletListAction() const;
    QAction* numberListAction() const;
    QAction* indentMoreAction() const;
    QAction* indentLessAction() const;
    QAction* lineSpacingAction() const;
    QAction* leftToRightAction() const;
    QAction* rightToLeftAction() const;

    // Insert actions
    QAction* insertTableAction() const;
    QAction* insertImageAction() const;
    QAction* insertShapeAction() const;
    QAction* insertChartAction() const;
    QAction* insertLinkAction() const;
    QAction* insertSymbolAction() const;
    QAction* insertHorizontalLineAction() const;
    QAction* insertDateAction() const;
    QAction* insertTimeAction() const;
    QAction* insertVideoAction() const;
    QAction* insertHeaderAction() const;
    QAction* insertFooterAction() const;

    // Page Layout actions
    QAction* pageSizeAction() const;
    QAction* pageMarginsAction() const;
    QAction* pageOrientationAction() const;
    QAction* pageBackgroundAction() const;

    // Review actions
    QAction* spellCheckAction() const;
    QAction* translateAction() const;
    QAction* wordCountAction() const;
    QAction* ttsAction() const;

    // Options action
    QAction* optionsAction() const;

    // View actions
    QAction* zoomInAction() const;
    QAction* zoomOutAction() const;
    QSlider* zoomSlider() const;
    QLabel* zoomLabel() const;
    QAction* showRulerAction() const;
    QAction* showStatusBarAction() const;

    // Help actions
    QAction* aboutAction() const;
    QAction* checkUpdatesAction() const;
    QAction* onlineHelpAction() const;
    QAction* websiteAction() const;
    QAction* pluginsAction() const;

    // Theme actions
    QAction* themeOffice2010Action() const;
    QAction* themeOffice2010SilverAction() const;
    QAction* themeOffice2010BlackAction() const;
    QAction* themeOffice2013Action() const;
    QAction* themeWindows8Action() const;

    // Menus
    QMenu* fileMenu() const;
    QMenu* editMenu() const;
    QMenu* insertMenu() const;
    QMenu* formatMenu() const;
    QMenu* viewMenu() const;
    QMenu* helpMenu() const;
    QMenu* recentFilesMenu() const;
    QMenu* themeMenu() const;
    QMenu* importMenu() const;
    QMenu* exportMenu() const;

    // Toolbars
    QToolBar* quickAccessToolBar() const;
    QToolBar* clipboardToolBar() const;
    QToolBar* fontToolBar() const;
    QToolBar* paragraphToolBar() const;
    QToolBar* editingToolBar() const;
    QToolBar* insertToolBar() const;
    QToolBar* pageLayoutToolBar() const;
    QToolBar* reviewToolBar() const;
    QToolBar* zoomToolBar() const;

    bool isUpdatingFont() const;
    void setUpdatingFont(bool updating);

private:
    static QIcon makeIcon(const QString& name);
    void setupIcons();

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
    QAction* m_exportPdfAction = nullptr;
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
    QAction* m_fontFaceAction = nullptr;
    QAction* m_fontSizeAction = nullptr;
    QAction* m_fontColorAction = nullptr;
    QAction* m_highlightColorAction = nullptr;
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
    QLabel* m_zoomLabel = nullptr;
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

    bool m_updatingFont = false;
};
