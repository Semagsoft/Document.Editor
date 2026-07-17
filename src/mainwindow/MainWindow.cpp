#include "MainWindow.h"
#include "StatusBarManager.h"
#include "ThemeManager.h"
#include "app/Application.h"
#include "app/Settings.h"
#include "editor/DocumentEditor.h"
#include "editor/DocumentManager.h"
#include "editor/DocumentTab.h"
#include "services/DocumentService.h"

#include <QMdiArea>
#include <QClipboard>
#include <QCloseEvent>
#include <QColorDialog>
#include <QComboBox>
#include <QDesktopServices>
#include <QFontDatabase>
#include <QImage>
#include <QInputDialog>
#include <QKeyEvent>
#include <QLabel>
#include <QMenuBar>
#include <QMessageBox>
#include <QSlider>
#include <QStatusBar>
#include <QTimer>
#include <QStyle>
#include <QToolBar>
#include <QUrl>

#include "dialogs/AboutDialog.h"
#include "dialogs/FindDialog.h"
#include "dialogs/GoToLineDialog.h"
#include "dialogs/InsertChartDialog.h"
#include "dialogs/InsertDateDialog.h"
#include "dialogs/InsertImageDialog.h"
#include "dialogs/InsertLineDialog.h"
#include "dialogs/InsertLinkDialog.h"
#include "dialogs/InsertShapeDialog.h"
#include "dialogs/InsertSymbolDialog.h"
#include "dialogs/InsertTableDialog.h"
#include "dialogs/InsertTimeDialog.h"
#include "dialogs/InsertVideoDialog.h"
#include "dialogs/LineSpacingDialog.h"
#include "dialogs/OptionsDialog.h"
#include "dialogs/ReplaceDialog.h"
#include "dialogs/SpellCheckDialog.h"

#include "plugins/PluginContext.h"
#include "plugins/PluginManager.h"
#include "utils/NetworkUtils.h"
#include "utils/SpellChecker.h"
#include "utils/TextUtils.h"
#include <QTextToSpeech>

static QIcon icon(const QString& name)
{
    return QIcon(QStringLiteral(":/icons/%1.svg").arg(name));
}

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
{
    m_settings = qobject_cast<Application*>(QApplication::instance())->settings();

    m_themeManager = new ThemeManager(this);

    setWindowTitle(QStringLiteral("Document.Editor"));
    setMinimumSize(800, 600);
    resize(1200, 800);

    createCentralArea();
    createMenuBar();
    createToolBars();
    createStatusBar();

    m_docManager = new DocumentManager(m_mdiArea, this, this);

    m_docService = new DocumentService(m_docManager, m_settings, m_statusBarManager, this);

    connect(m_docService, &DocumentService::documentOpened, this, [this](const QString &path) {
        if (!path.isEmpty())
            m_settings->addRecentFile(path);
        rebuildRecentFilesMenu();
        updateEditorActions();
        if (auto *e = currentEditor())
            e->setSpellChecker(m_spellChecker);
    });
    connect(m_docService, &DocumentService::statusMessage, this, [this](const QString &msg, int timeout) {
        statusBar()->showMessage(msg, timeout);
    });

    connectActions();
    loadSettings();

    // Startup dialog
    if (m_settings->showStartupDialog()) {
        QStringList recent = m_settings->recentFiles();
        if (!recent.isEmpty()) {
            statusBar()->showMessage(
                tr("Welcome back! Recent files are available under File > Recent Files"), 5000);
        }
    }

    // Check for updates on startup
    if (m_settings->checkForUpdatesOnStartup()) {
        auto* netUtils = new NetworkUtils(this);
        connect(netUtils, &NetworkUtils::updateAvailable, this, [this, netUtils](const QString& version, const QString& url) {
            statusBar()->showMessage(
                tr("Version %1 is available. Go to Help > Check for Updates.").arg(version), 8000);
            netUtils->deleteLater();
        });
        connect(netUtils, &NetworkUtils::upToDate, netUtils, &QObject::deleteLater);
        connect(netUtils, &NetworkUtils::updateCheckError, netUtils, &QObject::deleteLater);
        netUtils->checkForUpdates(QApplication::applicationVersion());
    }

    // Plugin system
    m_pluginManager = new PluginManager(this);
    m_pluginContext = new PluginContext(this);
    m_pluginContext->setMainWindow(this);
    m_pluginContext->setDocumentManager(m_docManager);
    m_pluginContext->setSettings(m_settings);
    m_pluginManager->setContext(m_pluginContext);
    m_pluginContext->setPluginManager(m_pluginManager);
    if (m_settings->pluginsEnabled())
        m_pluginManager->loadPlugins();

    // Spell checker
    m_spellChecker = new SpellChecker();

    // Text-to-speech
    m_tts = new QTextToSpeech(this);

    // Zoom debounce timer
    m_zoomDebounceTimer = new QTimer(this);
    m_zoomDebounceTimer->setSingleShot(true);
    m_zoomDebounceTimer->setInterval(50);
}

MainWindow::~MainWindow()
{
    if (m_pluginManager)
        m_pluginManager->unloadPlugins();
}

Settings* MainWindow::settings() const
{
    return m_settings;
}

StatusBarManager* MainWindow::statusBarManager() const
{
    return m_statusBarManager;
}

DocumentEditor* MainWindow::currentEditor() const
{
    return m_docManager->activeEditor();
}

void MainWindow::createCentralArea()
{
    m_mdiArea = new QMdiArea(this);
    m_mdiArea->setViewMode(QMdiArea::TabbedView);
    m_mdiArea->setTabsClosable(true);
    m_mdiArea->setTabsMovable(true);
    m_mdiArea->setDocumentMode(true);

    applyTabSettings();

    connect(m_mdiArea, &QMdiArea::subWindowActivated, this, [this]() {
        updateEditorActions();
    });
    setCentralWidget(m_mdiArea);
}

void MainWindow::applyTabSettings()
{
    m_mdiArea->setTabPosition(static_cast<QTabWidget::TabPosition>(m_settings->tabPlacement()));
    m_mdiArea->setTabShape(m_settings->tabSizeMode() == 0
            ? QTabWidget::Rounded
            : QTabWidget::Triangular);
}

void MainWindow::createMenuBar()
{
    // File menu
    m_fileMenu = menuBar()->addMenu(tr("&File"));

    m_newAction = m_fileMenu->addAction(tr("&New"));
    m_newAction->setShortcut(QKeySequence::New);
    m_newAction->setIcon(style()->standardIcon(QStyle::SP_FileIcon));

    m_openAction = m_fileMenu->addAction(tr("&Open..."));
    m_openAction->setShortcut(QKeySequence::Open);
    m_openAction->setIcon(style()->standardIcon(QStyle::SP_DialogOpenButton));

    m_fileMenu->addSeparator();

    m_closeAction = m_fileMenu->addAction(tr("&Close"));
    m_closeAction->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_W));
    m_closeAllButThisAction = m_fileMenu->addAction(tr("Close All But This"));
    m_closeAllAction = m_fileMenu->addAction(tr("Close All"));

    m_fileMenu->addSeparator();

    m_saveAction = m_fileMenu->addAction(tr("&Save"));
    m_saveAction->setShortcut(QKeySequence::Save);
    m_saveAction->setIcon(style()->standardIcon(QStyle::SP_DialogSaveButton));
    m_saveAsAction = m_fileMenu->addAction(tr("Save &As..."));
    m_saveAsAction->setShortcut(QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_S));
    m_saveCopyAction = m_fileMenu->addAction(tr("Save Cop&y..."));
    m_saveAllAction = m_fileMenu->addAction(tr("Save A&ll"));
    m_revertAction = m_fileMenu->addAction(tr("Re&vert"));

    m_fileMenu->addSeparator();

    m_importMenu = m_fileMenu->addMenu(tr("&Import"));
    m_importFtpAction = m_importMenu->addAction(tr("From &FTP..."));
    m_importArchiveAction = m_importMenu->addAction(tr("From &Archive..."));
    m_importImageAction = m_importMenu->addAction(tr("&Image..."));

    m_exportMenu = m_fileMenu->addMenu(tr("E&xport"));
    m_exportWordpressAction = m_exportMenu->addAction(tr("To &Wordpress..."));
    m_exportEmailAction = m_exportMenu->addAction(tr("To &Email..."));
    m_exportFtpAction = m_exportMenu->addAction(tr("To &FTP..."));
    m_exportMenu->addSeparator();
    m_exportXpsAction = m_exportMenu->addAction(tr("To &PDF..."));
    m_exportArchiveAction = m_exportMenu->addAction(tr("To &Archive..."));
    m_exportImageAction = m_exportMenu->addAction(tr("To &Image..."));
    m_exportSoundAction = m_exportMenu->addAction(tr("To &Sound..."));

    m_fileMenu->addSeparator();

    m_recentFilesMenu = m_fileMenu->addMenu(tr("&Recent Files"));

    m_fileMenu->addSeparator();

    m_optionsAction = m_fileMenu->addAction(tr("&Options..."));
    m_optionsAction->setMenuRole(QAction::PreferencesRole);

    m_fileMenu->addSeparator();

    m_printAction = m_fileMenu->addAction(tr("&Print..."));
    m_printAction->setShortcut(QKeySequence::Print);
    m_pageSetupAction = m_fileMenu->addAction(tr("Page Set&up..."));

    m_fileMenu->addSeparator();

    m_exitAction = m_fileMenu->addAction(tr("E&xit"));
    m_exitAction->setShortcut(QKeySequence::Quit);

    // Edit menu
    m_editMenu = menuBar()->addMenu(tr("&Edit"));

    m_undoAction = m_editMenu->addAction(tr("&Undo"));
    m_undoAction->setShortcut(QKeySequence::Undo);
    m_redoAction = m_editMenu->addAction(tr("&Redo"));
    m_redoAction->setShortcut(QKeySequence::Redo);

    m_editMenu->addSeparator();

    m_cutAction = m_editMenu->addAction(tr("Cu&t"));
    m_cutAction->setShortcut(QKeySequence::Cut);
    m_copyAction = m_editMenu->addAction(tr("&Copy"));
    m_copyAction->setShortcut(QKeySequence::Copy);
    m_pasteAction = m_editMenu->addAction(tr("&Paste"));
    m_pasteAction->setShortcut(QKeySequence::Paste);
    m_pasteTextAction = m_editMenu->addAction(tr("Paste &Text"));
    m_pasteImageAction = m_editMenu->addAction(tr("Paste &Image"));
    m_deleteAction = m_editMenu->addAction(tr("&Delete"));
    m_deleteAction->setShortcut(QKeySequence::Delete);

    m_editMenu->addSeparator();

    m_selectAllAction = m_editMenu->addAction(tr("Select &All"));
    m_selectAllAction->setShortcut(QKeySequence::SelectAll);

    m_editMenu->addSeparator();

    m_findAction = m_editMenu->addAction(tr("&Find..."));
    m_findAction->setShortcut(QKeySequence::Find);
    m_replaceAction = m_editMenu->addAction(tr("&Replace..."));
    m_replaceAction->setShortcut(QKeySequence::Replace);
    m_goToAction = m_editMenu->addAction(tr("&Go To..."));
    m_goToAction->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_G));

    // Insert menu
    m_insertMenu = menuBar()->addMenu(tr("&Insert"));

    m_insertTableAction = m_insertMenu->addAction(tr("&Table..."));
    m_insertMenu->addSeparator();
    m_insertImageAction = m_insertMenu->addAction(tr("&Image..."));
    m_insertShapeAction = m_insertMenu->addAction(tr("S&hape..."));
    m_insertChartAction = m_insertMenu->addAction(tr("&Chart..."));
    m_insertMenu->addSeparator();
    m_insertLinkAction = m_insertMenu->addAction(tr("&Link..."));
    m_insertSymbolAction = m_insertMenu->addAction(tr("S&ymbol..."));
    m_insertHorizontalLineAction = m_insertMenu->addAction(tr("Horizontal &Line..."));
    m_insertDateAction = m_insertMenu->addAction(tr("&Date..."));
    m_insertTimeAction = m_insertMenu->addAction(tr("&Time..."));
    m_insertVideoAction = m_insertMenu->addAction(tr("&Video..."));
    m_insertHeaderAction = m_insertMenu->addAction(tr("&Header..."));
    m_insertFooterAction = m_insertMenu->addAction(tr("&Footer..."));

    // Format menu
    m_formatMenu = menuBar()->addMenu(tr("F&ormat"));

    m_clearFormattingAction = m_formatMenu->addAction(tr("&Clear Formatting"));

    QMenu* fontSubMenu = m_formatMenu->addMenu(tr("&Font"));
    fontSubMenu->addAction(tr("Font &Face..."));
    fontSubMenu->addAction(tr("Font &Size..."));
    fontSubMenu->addAction(tr("Font &Color..."));
    fontSubMenu->addAction(tr("&Highlight Color..."));

    QMenu* styleSubMenu = m_formatMenu->addMenu(tr("S&tyle"));
    m_boldAction = styleSubMenu->addAction(tr("&Bold"));
    m_boldAction->setShortcut(QKeySequence::Bold);
    m_boldAction->setCheckable(true);
    m_italicAction = styleSubMenu->addAction(tr("&Italic"));
    m_italicAction->setShortcut(QKeySequence::Italic);
    m_italicAction->setCheckable(true);
    m_underlineAction = styleSubMenu->addAction(tr("&Underline"));
    m_underlineAction->setShortcut(QKeySequence::Underline);
    m_underlineAction->setCheckable(true);
    m_strikethroughAction = styleSubMenu->addAction(tr("&Strikethrough"));
    m_strikethroughAction->setCheckable(true);

    m_subscriptAction = m_formatMenu->addAction(tr("&Subscript"));
    m_superscriptAction = m_formatMenu->addAction(tr("Su&perscript"));
    m_formatMenu->addSeparator();
    m_indentMoreAction = m_formatMenu->addAction(tr("Indent &More"));
    m_indentLessAction = m_formatMenu->addAction(tr("Indent &Less"));

    QMenu* listSubMenu = m_formatMenu->addMenu(tr("&List"));
    m_bulletListAction = listSubMenu->addAction(tr("&Bullet List"));
    m_numberListAction = listSubMenu->addAction(tr("&Number List"));

    QMenu* alignSubMenu = m_formatMenu->addMenu(tr("Ali&gn"));
    m_alignLeftAction = alignSubMenu->addAction(tr("&Left"));
    m_alignCenterAction = alignSubMenu->addAction(tr("&Center"));
    m_alignRightAction = alignSubMenu->addAction(tr("&Right"));
    m_alignJustifyAction = alignSubMenu->addAction(tr("&Justify"));

    m_lineSpacingAction = m_formatMenu->addAction(tr("&Line Spacing..."));
    m_formatMenu->addSeparator();
    m_leftToRightAction = m_formatMenu->addAction(tr("&Left to Right"));
    m_rightToLeftAction = m_formatMenu->addAction(tr("&Right to Left"));

    QMenu* pageLayoutMenu = m_formatMenu->addMenu(tr("&Page Layout"));
    m_pageSizeAction = pageLayoutMenu->addAction(tr("Page &Size..."));
    m_pageMarginsAction = pageLayoutMenu->addAction(tr("&Margins..."));
    m_pageOrientationAction = pageLayoutMenu->addAction(tr("&Orientation"));
    m_pageBackgroundAction = pageLayoutMenu->addAction(tr("&Background..."));

    // View menu
    m_viewMenu = menuBar()->addMenu(tr("&View"));

    m_zoomInAction = m_viewMenu->addAction(tr("Zoom &In"));
    m_zoomInAction->setShortcut(QKeySequence::ZoomIn);
    m_zoomOutAction = m_viewMenu->addAction(tr("Zoom &Out"));
    m_zoomOutAction->setShortcut(QKeySequence::ZoomOut);
    m_viewMenu->addSeparator();
    m_showRulerAction = m_viewMenu->addAction(tr("Show &Ruler"));
    m_showRulerAction->setCheckable(true);
    m_showRulerAction->setChecked(m_settings->showRuler());
    m_showStatusBarAction = m_viewMenu->addAction(tr("Show &Status Bar"));
    m_showStatusBarAction->setCheckable(true);
    m_showStatusBarAction->setChecked(true);
    m_viewMenu->addSeparator();

    m_themeMenu = m_viewMenu->addMenu(tr("&Theme"));
    m_themeOffice2010Action = m_themeMenu->addAction(tr("Office 2010 &Blue"));
    m_themeOffice2010SilverAction = m_themeMenu->addAction(tr("Office 2010 &Silver"));
    m_themeOffice2010BlackAction = m_themeMenu->addAction(tr("Office 2010 &Black"));
    m_themeOffice2013Action = m_themeMenu->addAction(tr("Office &2013"));
    m_themeWindows8Action = m_themeMenu->addAction(tr("&Windows 8"));

    // Help menu
    m_helpMenu = menuBar()->addMenu(tr("&Help"));

    m_onlineHelpAction = m_helpMenu->addAction(tr("&Online Help..."));
    m_websiteAction = m_helpMenu->addAction(tr("&Website..."));
    m_helpMenu->addSeparator();
    m_checkUpdatesAction = m_helpMenu->addAction(tr("&Check for Updates..."));
    m_helpMenu->addSeparator();
    m_pluginsAction = m_helpMenu->addAction(tr("&Plugins..."));
    m_helpMenu->addSeparator();
    m_aboutAction = m_helpMenu->addAction(tr("&About Document.Editor..."));

    // ---- Assign SVG icons to all actions ----
    // File
    m_newAction->setIcon(icon(QStringLiteral("new")));
    m_openAction->setIcon(icon(QStringLiteral("open")));
    m_closeAction->setIcon(icon(QStringLiteral("close")));
    m_closeAllAction->setIcon(icon(QStringLiteral("close")));
    m_closeAllButThisAction->setIcon(icon(QStringLiteral("close")));
    m_saveAction->setIcon(icon(QStringLiteral("save")));
    m_saveAsAction->setIcon(icon(QStringLiteral("saveas")));
    m_saveCopyAction->setIcon(icon(QStringLiteral("saveas")));
    m_saveAllAction->setIcon(icon(QStringLiteral("save")));
    m_revertAction->setIcon(icon(QStringLiteral("undo")));
    m_printAction->setIcon(icon(QStringLiteral("print")));
    m_pageSetupAction->setIcon(icon(QStringLiteral("pagesize")));
    m_exitAction->setIcon(icon(QStringLiteral("export")));
    m_optionsAction->setIcon(icon(QStringLiteral("options")));
    m_importFtpAction->setIcon(icon(QStringLiteral("ftp")));
    m_importArchiveAction->setIcon(icon(QStringLiteral("archive")));
    m_importImageAction->setIcon(icon(QStringLiteral("image")));
    m_exportWordpressAction->setIcon(icon(QStringLiteral("export")));
    m_exportEmailAction->setIcon(icon(QStringLiteral("email")));
    m_exportFtpAction->setIcon(icon(QStringLiteral("ftp")));
    m_exportXpsAction->setIcon(icon(QStringLiteral("xps")));
    m_exportArchiveAction->setIcon(icon(QStringLiteral("archive")));
    m_exportImageAction->setIcon(icon(QStringLiteral("image")));
    m_exportSoundAction->setIcon(icon(QStringLiteral("sound")));

    // Edit
    m_undoAction->setIcon(icon(QStringLiteral("undo")));
    m_redoAction->setIcon(icon(QStringLiteral("redo")));
    m_cutAction->setIcon(icon(QStringLiteral("cut")));
    m_copyAction->setIcon(icon(QStringLiteral("copy")));
    m_pasteAction->setIcon(icon(QStringLiteral("paste")));
    m_pasteTextAction->setIcon(icon(QStringLiteral("paste")));
    m_pasteImageAction->setIcon(icon(QStringLiteral("image")));
    m_deleteAction->setIcon(icon(QStringLiteral("delete")));
    m_selectAllAction->setIcon(icon(QStringLiteral("selectall")));
    m_findAction->setIcon(icon(QStringLiteral("find")));
    m_replaceAction->setIcon(icon(QStringLiteral("replace")));
    m_goToAction->setIcon(icon(QStringLiteral("goto")));

    // Insert
    m_insertTableAction->setIcon(icon(QStringLiteral("table")));
    m_insertImageAction->setIcon(icon(QStringLiteral("image")));
    m_insertShapeAction->setIcon(icon(QStringLiteral("shape")));
    m_insertChartAction->setIcon(icon(QStringLiteral("chart")));
    m_insertLinkAction->setIcon(icon(QStringLiteral("link")));
    m_insertSymbolAction->setIcon(icon(QStringLiteral("symbol")));
    m_insertHorizontalLineAction->setIcon(icon(QStringLiteral("horizontalline")));
    m_insertDateAction->setIcon(icon(QStringLiteral("date")));
    m_insertTimeAction->setIcon(icon(QStringLiteral("time")));
    m_insertVideoAction->setIcon(icon(QStringLiteral("video")));
    m_insertHeaderAction->setIcon(icon(QStringLiteral("header")));
    m_insertFooterAction->setIcon(icon(QStringLiteral("footer")));

    // Format
    m_clearFormattingAction->setIcon(icon(QStringLiteral("clearformat")));
    m_boldAction->setIcon(icon(QStringLiteral("bold")));
    m_italicAction->setIcon(icon(QStringLiteral("italic")));
    m_underlineAction->setIcon(icon(QStringLiteral("underline")));
    m_strikethroughAction->setIcon(icon(QStringLiteral("strikethrough")));
    m_subscriptAction->setIcon(icon(QStringLiteral("subscript")));
    m_superscriptAction->setIcon(icon(QStringLiteral("superscript")));
    m_indentMoreAction->setIcon(icon(QStringLiteral("indentmore")));
    m_indentLessAction->setIcon(icon(QStringLiteral("indentless")));
    m_bulletListAction->setIcon(icon(QStringLiteral("bulletlist")));
    m_numberListAction->setIcon(icon(QStringLiteral("numberlist")));
    m_alignLeftAction->setIcon(icon(QStringLiteral("alignleft")));
    m_alignCenterAction->setIcon(icon(QStringLiteral("aligncenter")));
    m_alignRightAction->setIcon(icon(QStringLiteral("alignright")));
    m_alignJustifyAction->setIcon(icon(QStringLiteral("alignjustify")));
    m_lineSpacingAction->setIcon(icon(QStringLiteral("linespacing")));
    m_leftToRightAction->setIcon(icon(QStringLiteral("ltr")));
    m_rightToLeftAction->setIcon(icon(QStringLiteral("rtl")));
    m_pageSizeAction->setIcon(icon(QStringLiteral("pagesize")));
    m_pageMarginsAction->setIcon(icon(QStringLiteral("pagemargins")));
    m_pageOrientationAction->setIcon(icon(QStringLiteral("pageorientation")));
    m_pageBackgroundAction->setIcon(icon(QStringLiteral("pagebackground")));

    // View
    m_zoomInAction->setIcon(icon(QStringLiteral("zoomin")));
    m_zoomOutAction->setIcon(icon(QStringLiteral("zoomout")));
    m_showRulerAction->setIcon(icon(QStringLiteral("ruler")));
    m_showStatusBarAction->setIcon(icon(QStringLiteral("statusbar")));

    // Help
    m_onlineHelpAction->setIcon(icon(QStringLiteral("help")));
    m_websiteAction->setIcon(icon(QStringLiteral("website")));
    m_checkUpdatesAction->setIcon(icon(QStringLiteral("updates")));
    m_aboutAction->setIcon(icon(QStringLiteral("about")));
}

void MainWindow::createToolBars()
{
    m_quickAccessToolBar = addToolBar(tr("Quick Access"));
    m_quickAccessToolBar->setObjectName(QStringLiteral("QuickAccessToolBar"));
    m_quickAccessToolBar->setMovable(false);
    m_quickAccessToolBar->addAction(m_newAction);
    m_quickAccessToolBar->addAction(m_openAction);
    m_quickAccessToolBar->addAction(m_saveAction);
    m_quickAccessToolBar->addSeparator();
    m_quickAccessToolBar->addAction(m_undoAction);
    m_quickAccessToolBar->addAction(m_redoAction);

    m_clipboardToolBar = addToolBar(tr("Clipboard"));
    m_clipboardToolBar->setObjectName(QStringLiteral("ClipboardToolBar"));
    m_clipboardToolBar->addAction(m_cutAction);
    m_clipboardToolBar->addAction(m_copyAction);
    m_clipboardToolBar->addAction(m_pasteAction);

    m_fontToolBar = addToolBar(tr("Font"));
    m_fontToolBar->setObjectName(QStringLiteral("FontToolBar"));

    m_fontCombo = new QComboBox(this);
    m_fontCombo->setMinimumWidth(150);
    m_fontCombo->setEditable(false);
    const QFontDatabase fontDb;
    const QStringList families = fontDb.families();
    for (const QString& family : families)
        m_fontCombo->addItem(family);
    int defaultIdx = m_fontCombo->findText(m_settings->defaultFont().family());
    if (defaultIdx >= 0)
        m_fontCombo->setCurrentIndex(defaultIdx);
    m_fontToolBar->addWidget(m_fontCombo);

    m_fontSizeCombo = new QComboBox(this);
    m_fontSizeCombo->setEditable(true);
    m_fontSizeCombo->setMinimumWidth(60);
    const int sizes[] = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
    for (int size : sizes)
        m_fontSizeCombo->addItem(QString::number(size));
    m_fontSizeCombo->setCurrentText(QString::number(m_settings->defaultFontSize()));
    m_fontToolBar->addWidget(m_fontSizeCombo);

    m_fontToolBar->addSeparator();
    m_fontToolBar->addAction(m_boldAction);
    m_fontToolBar->addAction(m_italicAction);
    m_fontToolBar->addAction(m_underlineAction);
    m_fontToolBar->addAction(m_strikethroughAction);
    m_fontToolBar->addSeparator();
    m_fontToolBar->addAction(m_subscriptAction);
    m_fontToolBar->addAction(m_superscriptAction);

    m_paragraphToolBar = addToolBar(tr("Paragraph"));
    m_paragraphToolBar->setObjectName(QStringLiteral("ParagraphToolBar"));
    m_paragraphToolBar->addAction(m_alignLeftAction);
    m_paragraphToolBar->addAction(m_alignCenterAction);
    m_paragraphToolBar->addAction(m_alignRightAction);
    m_paragraphToolBar->addAction(m_alignJustifyAction);
    m_paragraphToolBar->addSeparator();
    m_paragraphToolBar->addAction(m_bulletListAction);
    m_paragraphToolBar->addAction(m_numberListAction);
    m_paragraphToolBar->addSeparator();
    m_paragraphToolBar->addAction(m_indentMoreAction);
    m_paragraphToolBar->addAction(m_indentLessAction);
    m_paragraphToolBar->addSeparator();
    m_paragraphToolBar->addAction(m_lineSpacingAction);

    m_editingToolBar = addToolBar(tr("Editing"));
    m_editingToolBar->setObjectName(QStringLiteral("EditingToolBar"));
    m_editingToolBar->addAction(m_findAction);
    m_editingToolBar->addAction(m_replaceAction);
    m_editingToolBar->addAction(m_goToAction);

    m_insertToolBar = addToolBar(tr("Insert"));
    m_insertToolBar->setObjectName(QStringLiteral("InsertToolBar"));
    m_insertToolBar->addAction(m_insertTableAction);
    m_insertToolBar->addAction(m_insertImageAction);
    m_insertToolBar->addAction(m_insertShapeAction);
    m_insertToolBar->addAction(m_insertChartAction);
    m_insertToolBar->addSeparator();
    m_insertToolBar->addAction(m_insertLinkAction);
    m_insertToolBar->addAction(m_insertSymbolAction);
    m_insertToolBar->addAction(m_insertHorizontalLineAction);
    m_insertToolBar->addSeparator();
    m_insertToolBar->addAction(m_insertDateAction);
    m_insertToolBar->addAction(m_insertTimeAction);
    m_insertToolBar->addAction(m_insertVideoAction);

    m_pageLayoutToolBar = addToolBar(tr("Page Layout"));
    m_pageLayoutToolBar->setObjectName(QStringLiteral("PageLayoutToolBar"));

    m_pageLayoutToolBar->addAction(m_pageSizeAction);
    m_pageLayoutToolBar->addAction(m_pageMarginsAction);
    m_pageLayoutToolBar->addAction(m_pageOrientationAction);
    m_pageLayoutToolBar->addAction(m_pageBackgroundAction);

    m_reviewToolBar = addToolBar(tr("Review"));
    m_reviewToolBar->setObjectName(QStringLiteral("ReviewToolBar"));
    m_spellCheckAction = m_reviewToolBar->addAction(tr("Spell Check..."));
    m_spellCheckAction->setIcon(icon(QStringLiteral("spellcheck")));
    m_translateAction = m_reviewToolBar->addAction(tr("Translate..."));
    m_translateAction->setIcon(icon(QStringLiteral("translate")));
    m_wordCountAction = m_reviewToolBar->addAction(tr("Word Count"));
    m_wordCountAction->setIcon(icon(QStringLiteral("wordcount")));
    m_reviewToolBar->addSeparator();
    m_ttsAction = m_reviewToolBar->addAction(tr("Speak"));
    m_ttsAction->setIcon(icon(QStringLiteral("sound")));

    m_zoomToolBar = addToolBar(tr("Zoom"));
    m_zoomToolBar->setObjectName(QStringLiteral("ZoomToolBar"));
    m_zoomToolBar->addAction(m_zoomOutAction);

    m_zoomSlider = new QSlider(Qt::Horizontal, this);
    m_zoomSlider->setRange(10, 500);
    m_zoomSlider->setValue(100);
    m_zoomSlider->setFixedWidth(120);
    m_zoomToolBar->addWidget(m_zoomSlider);

    m_zoomToolBar->addAction(m_zoomInAction);

    QLabel* zoomLabel = new QLabel(QStringLiteral(" 100%"), this);
    m_zoomToolBar->addWidget(zoomLabel);

    connect(m_zoomSlider, &QSlider::valueChanged, this,
        [this, zoomLabel](int value) {
            zoomLabel->setText(QStringLiteral(" %1%").arg(value));
            m_statusBarManager->setZoomLevel(value);
            m_zoomDebounceTimer->start();
        });
    connect(m_zoomDebounceTimer, &QTimer::timeout, this, [this]() {
        DocumentEditor* editor = currentEditor();
        if (editor)
            editor->setZoomLevel(m_zoomSlider->value() / 100.0);
    });
    connect(m_zoomInAction, &QAction::triggered, this, [this]() {
        int v = m_zoomSlider->value();
        m_zoomSlider->setValue(qMin(v + 10, m_zoomSlider->maximum()));
    });
    connect(m_zoomOutAction, &QAction::triggered, this, [this]() {
        int v = m_zoomSlider->value();
        m_zoomSlider->setValue(qMax(v - 10, m_zoomSlider->minimum()));
    });
}

void MainWindow::createStatusBar()
{
    m_statusBarManager = new StatusBarManager(statusBar(), this);
}

void MainWindow::connectActions()
{
    connectFileActions();
    connectEditActions();
    connectFormatActions();
    connectInsertActions();
    connectPageLayoutActions();
    connectReviewActions();
    connectViewActions();
    connectHelpActions();
    connectDocumentSignals();
}

void MainWindow::connectFileActions()
{
    connect(m_newAction, &QAction::triggered, m_docService, &DocumentService::newDocument);
    connect(m_openAction, &QAction::triggered, this, [this]() { m_docService->openDocument(); });
    connect(m_closeAction, &QAction::triggered, this, &MainWindow::closeCurrentDocument);
    connect(m_closeAllAction, &QAction::triggered, this, &MainWindow::closeAllDocuments);
    connect(m_closeAllButThisAction, &QAction::triggered, this, &MainWindow::closeAllButCurrent);
    connect(m_saveAction, &QAction::triggered, this, [this]() { m_docManager->saveDocument(); });
    connect(m_saveAsAction, &QAction::triggered, this, [this]() { m_docManager->saveDocumentAs(); });
    connect(m_saveCopyAction, &QAction::triggered, this, [this]() { m_docManager->saveDocumentCopy(); });
    connect(m_saveAllAction, &QAction::triggered, this, [this]() { m_docManager->saveAllDocuments(); });
    connect(m_exitAction, &QAction::triggered, this, &QWidget::close);

    connect(m_importFtpAction, &QAction::triggered, m_docService, &DocumentService::importFtp);
    connect(m_importArchiveAction, &QAction::triggered, m_docService, &DocumentService::importArchive);
    connect(m_importImageAction, &QAction::triggered, m_docService, &DocumentService::importImage);

    connect(m_exportWordpressAction, &QAction::triggered, m_docService, &DocumentService::exportWordpress);
    connect(m_exportEmailAction, &QAction::triggered, m_docService, &DocumentService::exportEmail);
    connect(m_exportFtpAction, &QAction::triggered, m_docService, &DocumentService::exportFtp);
    connect(m_exportXpsAction, &QAction::triggered, m_docService, &DocumentService::exportPdf);
    connect(m_exportArchiveAction, &QAction::triggered, m_docService, &DocumentService::exportArchive);
    connect(m_exportImageAction, &QAction::triggered, m_docService, &DocumentService::exportImage);
    connect(m_exportSoundAction, &QAction::triggered, m_docService, &DocumentService::exportSound);

    connect(m_printAction, &QAction::triggered, m_docService, &DocumentService::printDocument);
    connect(m_pageSetupAction, &QAction::triggered, m_docService, &DocumentService::pageSetup);

    connect(m_revertAction, &QAction::triggered, m_docService, &DocumentService::revertDocument);

    connect(m_optionsAction, &QAction::triggered, this, [this]() {
        OptionsDialog dlg(m_settings, this);
        if (dlg.exec() == QDialog::Accepted) {
            m_settings->save();
            rebuildRecentFilesMenu();
            emit documentChanged();
        }
    });
    connect(m_settings, &Settings::settingsChanged, this, &MainWindow::rebuildRecentFilesMenu);
}

void MainWindow::connectEditActions()
{
    connect(m_undoAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->undo(); });
    connect(m_redoAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->redo(); });
    connect(m_cutAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->cut(); });
    connect(m_copyAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->copy(); });
    connect(m_pasteAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->paste(); });
    connect(m_pasteTextAction, &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor()) {
            QString text = QApplication::clipboard()->text();
            if (!text.isEmpty())
                e->textCursor().insertText(text);
        }
    });
    connect(m_pasteImageAction, &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor()) {
            QImage img = QApplication::clipboard()->image();
            if (!img.isNull())
                e->textCursor().insertImage(img);
        }
    });
    connect(m_deleteAction, &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor()) {
            QTextCursor c = e->textCursor();
            c.removeSelectedText();
        }
    });
    connect(m_selectAllAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->selectAll(); });
    connect(m_findAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        FindDialog dlg(this);
        connect(&dlg, &FindDialog::findNext, this, [this](const QString& text, bool matchCase, bool wholeWord) {
            QTextDocument::FindFlags flags;
            if (matchCase)
                flags |= QTextDocument::FindCaseSensitively;
            if (wholeWord)
                flags |= QTextDocument::FindWholeWords;
            currentEditor()->find(text, flags);
        });
        connect(&dlg, &FindDialog::findPrevious, this, [this](const QString& text, bool matchCase, bool wholeWord) {
            QTextDocument::FindFlags flags = QTextDocument::FindBackward;
            if (matchCase)
                flags |= QTextDocument::FindCaseSensitively;
            if (wholeWord)
                flags |= QTextDocument::FindWholeWords;
            currentEditor()->find(text, flags);
        });
        dlg.exec();
    });
    connect(m_replaceAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        ReplaceDialog dlg(this);
        connect(&dlg, &ReplaceDialog::findNext, this, [this](const QString& text, bool matchCase) {
            QTextDocument::FindFlags flags;
            if (matchCase)
                flags |= QTextDocument::FindCaseSensitively;
            currentEditor()->find(text, flags);
        });
        connect(&dlg, &ReplaceDialog::replace, this, [this](const QString& find, const QString& replace, bool matchCase) {
            QTextDocument::FindFlags flags;
            if (matchCase)
                flags |= QTextDocument::FindCaseSensitively;
            if (currentEditor()->find(find, flags))
                currentEditor()->textCursor().insertText(replace);
        });
        connect(&dlg, &ReplaceDialog::replaceAll, this, [this](const QString& find, const QString& replace, bool matchCase) {
            QTextDocument::FindFlags flags;
            if (matchCase)
                flags |= QTextDocument::FindCaseSensitively;
            int count = 0;
            QTextCursor cursor = currentEditor()->textCursor();
            cursor.movePosition(QTextCursor::Start);
            while (!cursor.isNull() && cursor.position() < currentEditor()->document()->characterCount()) {
                cursor = currentEditor()->document()->find(find, cursor, flags);
                if (!cursor.isNull()) {
                    cursor.insertText(replace);
                    count++;
                }
            }
        });
        dlg.exec();
    });
    connect(m_goToAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        GoToLineDialog dlg(currentEditor()->selectedLineNumber(),
            currentEditor()->lineCount(), this);
        if (dlg.exec() == QDialog::Accepted)
            currentEditor()->goToLine(dlg.lineNumber());
    });
}

void MainWindow::connectFormatActions()
{
    connect(m_boldAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleBold(); });
    connect(m_italicAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleItalic(); });
    connect(m_underlineAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleUnderline(); });
    connect(m_strikethroughAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleStrikethrough(); });
    connect(m_subscriptAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleSubscript(); });
    connect(m_superscriptAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleSuperscript(); });
    connect(m_clearFormattingAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->clearFormatting(); });
    connect(m_alignLeftAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignLeft); });
    connect(m_alignCenterAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignCenter); });
    connect(m_alignRightAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignRight); });
    connect(m_alignJustifyAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignJustify); });
    connect(m_bulletListAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleBulletList(); });
    connect(m_numberListAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleNumberList(); });
    connect(m_indentMoreAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->indentMore(); });
    connect(m_indentLessAction, &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->indentLess(); });
    connect(m_lineSpacingAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        qreal current = currentEditor()->textCursor().blockFormat().lineHeight() / 100.0;
        LineSpacingDialog dlg(current, this);
        if (dlg.exec() == QDialog::Accepted)
            currentEditor()->setLineSpacing(dlg.lineSpacing());
    });
    connect(m_leftToRightAction, &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor())
            e->setTextDirection(Qt::LeftToRight);
    });
    connect(m_rightToLeftAction, &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor())
            e->setTextDirection(Qt::RightToLeft);
    });

    connect(m_fontCombo, &QComboBox::currentTextChanged, this, [this](const QString& family) {
        if (m_updatingFont || !currentEditor())
            return;
        QTextCharFormat fmt;
        fmt.setFontFamilies({ family });
        currentEditor()->textCursor().mergeCharFormat(fmt);
    });
    connect(m_fontSizeCombo, &QComboBox::currentTextChanged, this, [this](const QString& text) {
        if (m_updatingFont || !currentEditor())
            return;
        bool ok;
        qreal size = text.toDouble(&ok);
        if (ok && size > 0) {
            QTextCharFormat fmt;
            fmt.setFontPointSize(size);
            currentEditor()->textCursor().mergeCharFormat(fmt);
        }
    });
}

void MainWindow::connectInsertActions()
{
    connect(m_insertTableAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertTableDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QTextCursor cursor = currentEditor()->textCursor();
            cursor.insertTable(dlg.rows(), dlg.columns());
        }
    });
    connect(m_insertImageAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertImageDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            DocumentService::embedImageInDocument(currentEditor(), dlg.imagePath(), dlg.width(), dlg.height());
    });
    connect(m_insertShapeAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertShapeDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QPixmap shape = dlg.generateShape(dlg.shapeName());
            if (!shape.isNull()) {
                QTextImageFormat fmt;
                fmt.setName(QStringLiteral("shape_temp"));
                fmt.setWidth(shape.width());
                fmt.setHeight(shape.height());
                currentEditor()->textCursor().insertImage(shape.toImage());
            }
        }
    });
    connect(m_insertChartAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertChartDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QPixmap chart = dlg.generateChart(dlg.chartType());
            if (!chart.isNull())
                currentEditor()->textCursor().insertImage(chart.toImage());
        }
    });
    connect(m_insertLinkAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertLinkDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QTextCursor cursor = currentEditor()->textCursor();
            cursor.insertHtml(QStringLiteral("<a href=\"%1\">%2</a>").arg(dlg.url().toHtmlEscaped(), dlg.displayText().toHtmlEscaped()));
        }
    });
    connect(m_insertSymbolAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertSymbolDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            currentEditor()->textCursor().insertText(QString(dlg.selectedSymbol()));
    });
    connect(m_insertHorizontalLineAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertLineDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            currentEditor()->textCursor().insertHtml(QStringLiteral("<hr>"));
    });
    connect(m_insertDateAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertDateDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            currentEditor()->textCursor().insertText(
                QDate::currentDate().toString(dlg.dateFormat()));
    });
    connect(m_insertTimeAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertTimeDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            currentEditor()->textCursor().insertText(
                QTime::currentTime().toString(dlg.timeFormat()));
    });
    connect(m_insertVideoAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        InsertVideoDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QTextCursor cursor = currentEditor()->textCursor();
            cursor.insertHtml(QStringLiteral("<a href=\"%1\">%1</a>").arg(dlg.videoPath().toHtmlEscaped()));
        }
    });
    connect(m_insertHeaderAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        QTextCursor cursor = currentEditor()->textCursor();
        cursor.movePosition(QTextCursor::Start);
        cursor.insertHtml(QStringLiteral(
            "<div style=\"border-bottom: 2px solid #444; padding-bottom: 6px; "
            "margin-bottom: 12px; font-size: 10pt; color: #666;\">Header</div>"));
    });
    connect(m_insertFooterAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        QTextCursor cursor = currentEditor()->textCursor();
        cursor.movePosition(QTextCursor::End);
        cursor.insertHtml(QStringLiteral(
            "<div style=\"border-top: 2px solid #444; padding-top: 6px; "
            "margin-top: 12px; font-size: 10pt; color: #666;\">Footer</div>"));
    });
}

void MainWindow::connectPageLayoutActions()
{
    connect(m_pageSizeAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        qreal w = currentEditor()->pageWidth();
        qreal h = currentEditor()->pageHeight();
        bool ok;
        QString result = QInputDialog::getText(this, tr("Page Size"),
            tr("Width x Height (points):"), QLineEdit::Normal,
            QStringLiteral("%1 x %2").arg(w).arg(h), &ok);
        if (!ok)
            return;
        QStringList parts = result.split(QStringLiteral("x"), Qt::SkipEmptyParts);
        if (parts.size() == 2) {
            currentEditor()->setPageWidth(parts[0].trimmed().toDouble());
            currentEditor()->setPageHeight(parts[1].trimmed().toDouble());
        }
    });
    connect(m_pageMarginsAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        QMarginsF m = currentEditor()->pageMargins();
        bool ok;
        QString result = QInputDialog::getText(this, tr("Page Margins"),
            tr("Left, Top, Right, Bottom (points):"), QLineEdit::Normal,
            QStringLiteral("%1, %2, %3, %4").arg(m.left()).arg(m.top()).arg(m.right()).arg(m.bottom()), &ok);
        if (!ok)
            return;
        QStringList parts = result.split(QStringLiteral(","), Qt::SkipEmptyParts);
        if (parts.size() == 4) {
            currentEditor()->setPageMargins(QMarginsF(
                parts[0].trimmed().toDouble(), parts[1].trimmed().toDouble(),
                parts[2].trimmed().toDouble(), parts[3].trimmed().toDouble()));
        }
    });
    connect(m_pageOrientationAction, &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor())
            e->togglePageOrientation();
    });
    connect(m_pageBackgroundAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        QColor color = QColorDialog::getColor(currentEditor()->pageBackground(), this,
            tr("Page Background"));
        if (color.isValid())
            currentEditor()->setPageBackground(color);
    });
}

void MainWindow::connectReviewActions()
{
    connect(m_spellCheckAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        if (!m_settings->spellCheckEnabled()) {
            QMessageBox::information(this, tr("Spell Check"),
                tr("Spell check is disabled. Enable it in Options > Editing."));
            return;
        }
        DocumentEditor* e = currentEditor();
        if (!m_spellChecker->hasDictionary()) {
            QMessageBox::information(this, tr("Spell Check"),
                tr("Spell check is unavailable because no system dictionary was found.\n\n"
                   "Install a dictionary (e.g., /usr/share/dict/words) to enable spell checking."));
            return;
        }
        QString text = e->toPlainText();
        QStringList misspelled = m_spellChecker->findMisspelled(text);
        if (misspelled.isEmpty()) {
            QMessageBox::information(this, tr("Spell Check"), tr("No misspellings found."));
        } else {
            SpellCheckDialog dlg(m_spellChecker, misspelled, this);
            if (dlg.exec() == QDialog::Accepted) {
                const auto reps = dlg.replacements();
                for (const auto& rep : reps) {
                    if (rep.original.isEmpty())
                        continue;
                    QTextCursor cursor(e->document());
                    cursor.movePosition(QTextCursor::Start);
                    cursor = e->document()->find(rep.original, cursor);
                    while (!cursor.isNull()) {
                        cursor.insertText(rep.replacement);
                        int pos = cursor.position();
                        cursor = e->document()->find(rep.original, cursor);
                        if (!cursor.isNull() && cursor.position() == pos)
                            break;
                    }
                }
            }
        }
    });
    connect(m_translateAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        QString text = currentEditor()->textCursor().selectedText();
        if (text.isEmpty())
            text = currentEditor()->toPlainText().left(500);
        if (text.isEmpty())
            return;
        QDesktopServices::openUrl(QUrl(QStringLiteral("https://translate.google.com/?text=")
            + QUrl::toPercentEncoding(text)));
    });
    connect(m_wordCountAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        DocumentEditor* e = currentEditor();
        int wc = e->wordCount();
        int cc = e->toPlainText().size();
        int selected = e->textCursor().hasSelection()
            ? e->textCursor().selectedText().size()
            : 0;
        QMessageBox::information(this, tr("Word Count"),
            tr("Words: %1\nCharacters: %2\nSelected characters: %3").arg(wc).arg(cc).arg(selected));
    });
    connect(m_ttsAction, &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        if (m_tts->state() == QTextToSpeech::Speaking) {
            m_tts->stop();
            m_ttsAction->setText(tr("Speak"));
            return;
        }
        QString text = currentEditor()->textCursor().selectedText();
        if (text.isEmpty())
            text = currentEditor()->toPlainText();
        if (!text.isEmpty()) {
            m_tts->say(text);
            m_ttsAction->setText(tr("Stop"));
        }
    });
    connect(m_tts, &QTextToSpeech::stateChanged, this, [this](QTextToSpeech::State state) {
        if (state == QTextToSpeech::Ready || state == QTextToSpeech::Error)
            m_ttsAction->setText(tr("Speak"));
    });
}

void MainWindow::connectViewActions()
{
    connect(m_showRulerAction, &QAction::toggled, this, [this](bool checked) {
        m_settings->setShowRuler(checked);
    });
    connect(m_showStatusBarAction, &QAction::toggled, this, [this](bool checked) {
        statusBar()->setVisible(checked);
    });
    connect(m_themeOffice2010Action, &QAction::triggered, this, [this]() {
        m_themeManager->setTheme(ThemeManager::Office2010Blue);
        m_settings->setTheme(static_cast<int>(ThemeManager::Office2010Blue));
    });
    connect(m_themeOffice2010SilverAction, &QAction::triggered, this, [this]() {
        m_themeManager->setTheme(ThemeManager::Office2010Silver);
        m_settings->setTheme(static_cast<int>(ThemeManager::Office2010Silver));
    });
    connect(m_themeOffice2010BlackAction, &QAction::triggered, this, [this]() {
        m_themeManager->setTheme(ThemeManager::Office2010Black);
        m_settings->setTheme(static_cast<int>(ThemeManager::Office2010Black));
    });
    connect(m_themeOffice2013Action, &QAction::triggered, this, [this]() {
        m_themeManager->setTheme(ThemeManager::Office2013);
        m_settings->setTheme(static_cast<int>(ThemeManager::Office2013));
    });
    connect(m_themeWindows8Action, &QAction::triggered, this, [this]() {
        m_themeManager->setTheme(ThemeManager::Windows8);
        m_settings->setTheme(static_cast<int>(ThemeManager::Windows8));
    });
}

void MainWindow::connectHelpActions()
{
    connect(m_aboutAction, &QAction::triggered, this, [this]() {
        AboutDialog dlg(this);
        dlg.exec();
    });
    connect(m_onlineHelpAction, &QAction::triggered, this, []() {
        QDesktopServices::openUrl(QUrl(QStringLiteral("http://documenteditor.net/documentation/")));
    });
    connect(m_websiteAction, &QAction::triggered, this, []() {
        QDesktopServices::openUrl(QUrl(QStringLiteral("http://documenteditor.net")));
    });
    connect(m_checkUpdatesAction, &QAction::triggered, this, [this]() {
        auto* netUtils = new NetworkUtils(this);
        connect(netUtils, &NetworkUtils::updateAvailable, this, [this, netUtils](const QString& version, const QString& url) {
            QMessageBox::information(this, tr("Update Available"),
                tr("Version %1 is available.\n%2").arg(version, url));
            netUtils->deleteLater();
        });
        connect(netUtils, &NetworkUtils::upToDate, this, [this, netUtils]() {
            QMessageBox::information(this, tr("No Updates"),
                tr("You are using the latest version."));
            netUtils->deleteLater();
        });
        connect(netUtils, &NetworkUtils::updateCheckError, this, [this, netUtils](const QString& error) {
            QMessageBox::warning(this, tr("Update Check Failed"),
                tr("Could not check for updates:\n%1").arg(error));
            netUtils->deleteLater();
        });
        netUtils->checkForUpdates(QApplication::applicationVersion());
    });
    connect(m_pluginsAction, &QAction::triggered, this, [this]() {
        QStringList names = m_pluginManager->loadedPluginNames();
        if (names.isEmpty()) {
            QMessageBox::information(this, tr("Plugins"),
                tr("No plugins are currently loaded.\n\n"
                   "Place plugin DLLs/shared libraries in the plugins folder\n"
                   "and enable plugins in Options to load them."));
        } else {
            QMessageBox::information(this, tr("Plugins"),
                tr("Loaded plugins (%1):\n%2").arg(names.size()).arg(names.join(QStringLiteral("\n"))));
        }
    });
}

void MainWindow::connectDocumentSignals()
{
    connect(m_docManager, &DocumentManager::activeTabChanged, this, [this](DocumentTab* tab) {
        updateEditorActions();
        if (tab) {
            connectTabInsertSignals(tab);
            if (auto *e = tab->editor())
                e->setSpellChecker(m_spellChecker);
        }
    });
    connect(m_docManager, &DocumentManager::statusLineColumnChanged,
        m_statusBarManager, &StatusBarManager::setLineColumn);
    connect(m_docManager, &DocumentManager::statusWordCountChanged,
        m_statusBarManager, &StatusBarManager::setWordCount);
}

void MainWindow::updateEditorActions()
{
    DocumentEditor* editor = currentEditor();
    bool hasDoc = editor != nullptr;

    m_saveAction->setEnabled(hasDoc);
    m_saveAsAction->setEnabled(hasDoc);
    m_saveCopyAction->setEnabled(hasDoc);
    m_saveAllAction->setEnabled(hasDoc && m_docManager->tabCount() > 1);
    m_closeAction->setEnabled(hasDoc);
    m_closeAllAction->setEnabled(hasDoc && m_docManager->tabCount() > 1);
    m_closeAllButThisAction->setEnabled(hasDoc && m_docManager->tabCount() > 1);
    m_optionsAction->setEnabled(true);
    m_revertAction->setEnabled(hasDoc && !editor->documentName().isEmpty());
    m_printAction->setEnabled(hasDoc);
    m_pageSetupAction->setEnabled(hasDoc);

    m_undoAction->setEnabled(hasDoc && editor->document()->isUndoAvailable());
    m_redoAction->setEnabled(hasDoc && editor->document()->isRedoAvailable());
    m_cutAction->setEnabled(hasDoc && editor->textCursor().hasSelection());
    m_copyAction->setEnabled(hasDoc && editor->textCursor().hasSelection());
    m_pasteAction->setEnabled(hasDoc);
    m_pasteTextAction->setEnabled(hasDoc);
    m_pasteImageAction->setEnabled(hasDoc);
    m_deleteAction->setEnabled(hasDoc && editor->textCursor().hasSelection());
    m_selectAllAction->setEnabled(hasDoc);
    m_findAction->setEnabled(hasDoc);
    m_replaceAction->setEnabled(hasDoc);
    m_goToAction->setEnabled(hasDoc);

    m_boldAction->setEnabled(hasDoc);
    m_italicAction->setEnabled(hasDoc);
    m_underlineAction->setEnabled(hasDoc);
    m_strikethroughAction->setEnabled(hasDoc);
    m_subscriptAction->setEnabled(hasDoc);
    m_superscriptAction->setEnabled(hasDoc);
    m_clearFormattingAction->setEnabled(hasDoc);
    m_fontCombo->setEnabled(hasDoc);
    m_fontSizeCombo->setEnabled(hasDoc);

    m_alignLeftAction->setEnabled(hasDoc);
    m_alignCenterAction->setEnabled(hasDoc);
    m_alignRightAction->setEnabled(hasDoc);
    m_alignJustifyAction->setEnabled(hasDoc);
    m_bulletListAction->setEnabled(hasDoc);
    m_numberListAction->setEnabled(hasDoc);
    m_indentMoreAction->setEnabled(hasDoc);
    m_indentLessAction->setEnabled(hasDoc);
    m_lineSpacingAction->setEnabled(hasDoc);
    m_leftToRightAction->setEnabled(hasDoc);
    m_rightToLeftAction->setEnabled(hasDoc);

    m_insertTableAction->setEnabled(hasDoc);
    m_insertImageAction->setEnabled(hasDoc);
    m_insertShapeAction->setEnabled(hasDoc);
    m_insertChartAction->setEnabled(hasDoc);
    m_insertLinkAction->setEnabled(hasDoc);
    m_insertSymbolAction->setEnabled(hasDoc);
    m_insertHorizontalLineAction->setEnabled(hasDoc);
    m_insertDateAction->setEnabled(hasDoc);
    m_insertTimeAction->setEnabled(hasDoc);
    m_insertVideoAction->setEnabled(hasDoc);
    m_insertHeaderAction->setEnabled(hasDoc);
    m_insertFooterAction->setEnabled(hasDoc);

    m_pageSizeAction->setEnabled(hasDoc);
    m_pageMarginsAction->setEnabled(hasDoc);
    m_pageOrientationAction->setEnabled(hasDoc);
    m_pageBackgroundAction->setEnabled(hasDoc);

    m_spellCheckAction->setEnabled(hasDoc);
    m_translateAction->setEnabled(hasDoc);
    m_wordCountAction->setEnabled(hasDoc);
    m_ttsAction->setEnabled(hasDoc);

    m_zoomSlider->setEnabled(hasDoc);
    m_zoomInAction->setEnabled(hasDoc);
    m_zoomOutAction->setEnabled(hasDoc);

    if (hasDoc) {
        QTextCharFormat fmt = editor->textCursor().charFormat();

        m_updatingFont = true;
        const QFont& f = fmt.font();
        int idx = m_fontCombo->findText(f.family());
        if (idx >= 0)
            m_fontCombo->setCurrentIndex(idx);
        m_fontSizeCombo->setCurrentText(QString::number(qRound(f.pointSizeF())));
        m_updatingFont = false;

        m_boldAction->setChecked(fmt.fontWeight() >= QFont::Bold);
        m_italicAction->setChecked(fmt.fontItalic());
        m_underlineAction->setChecked(fmt.fontUnderline());
        m_strikethroughAction->setChecked(fmt.fontStrikeOut());
    } else {
        m_statusBarManager->setLineColumn(0, 0, 0, 0);
        m_statusBarManager->setWordCount(0);
        m_statusBarManager->setFileSize(QStringLiteral("0 KB"));
    }
}

void MainWindow::connectTabInsertSignals(DocumentTab* tab)
{
    if (!tab)
        return;
    disconnect(tab, nullptr, this, nullptr);
    connect(tab, &DocumentTab::insertTableRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertTableDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            e->textCursor().insertTable(dlg.rows(), dlg.columns());
    });
    connect(tab, &DocumentTab::insertImageRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertImageDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            DocumentService::embedImageInDocument(e, dlg.imagePath(), dlg.width(), dlg.height());
    });
    connect(tab, &DocumentTab::insertShapeRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertShapeDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QPixmap shape = dlg.generateShape(dlg.shapeName());
            if (!shape.isNull())
                e->textCursor().insertImage(shape.toImage());
        }
    });
    connect(tab, &DocumentTab::insertChartRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertChartDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QPixmap chart = dlg.generateChart(dlg.chartType());
            if (!chart.isNull())
                e->textCursor().insertImage(chart.toImage());
        }
    });
    connect(tab, &DocumentTab::insertLinkRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertLinkDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QTextCursor cursor = e->textCursor();
            cursor.insertHtml(QStringLiteral("<a href=\"%1\">%2</a>")
                    .arg(dlg.url().toHtmlEscaped(), dlg.displayText().toHtmlEscaped()));
        }
    });
    connect(tab, &DocumentTab::insertSymbolRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertSymbolDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            e->textCursor().insertText(QString(dlg.selectedSymbol()));
    });
    connect(tab, &DocumentTab::insertDateRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertDateDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            e->textCursor().insertText(QDate::currentDate().toString(dlg.dateFormat()));
    });
    connect(tab, &DocumentTab::insertTimeRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertTimeDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted)
            e->textCursor().insertText(QTime::currentTime().toString(dlg.timeFormat()));
    });
    connect(tab, &DocumentTab::insertVideoRequested, this, [this, tab]() {
        DocumentEditor* e = tab->editor();
        if (!e)
            return;
        InsertVideoDialog dlg(this);
        if (dlg.exec() == QDialog::Accepted) {
            QTextCursor cursor = e->textCursor();
            cursor.insertHtml(QStringLiteral("<a href=\"%1\">%1</a>").arg(dlg.videoPath().toHtmlEscaped()));
        }
    });
}

void MainWindow::rebuildRecentFilesMenu()
{
    m_recentFilesMenu->clear();
    if (!m_settings->showRecentDocuments())
        return;
    QStringList recent = m_settings->recentFiles();
    if (recent.isEmpty()) {
        QAction* emptyAction = m_recentFilesMenu->addAction(tr("No recent files"));
        emptyAction->setEnabled(false);
        return;
    }
    for (const QString& file : recent) {
        QAction* action = m_recentFilesMenu->addAction(file);
        connect(action, &QAction::triggered, this, [this, file]() {
            m_docService->openDocument(file);
        });
    }
    m_recentFilesMenu->addSeparator();
    QAction* clearAction = m_recentFilesMenu->addAction(tr("Clear Recent Files"));
    connect(clearAction, &QAction::triggered, this, [this]() {
        m_settings->setRecentFiles({ });
        rebuildRecentFilesMenu();
    });
}

void MainWindow::loadSettings()
{
    QRect geom = m_settings->windowGeometry();
    if (geom.isValid())
        setGeometry(geom);
    if (m_settings->windowMaximized())
        showMaximized();

    statusBar()->setVisible(m_showStatusBarAction->isChecked());
    applyTabSettings();

    if (m_settings->enableGlass())
        setAttribute(Qt::WA_TranslucentBackground);

    int themeIdx = m_settings->theme();
    if (themeIdx >= ThemeManager::Office2010Blue && themeIdx <= ThemeManager::Windows8)
        m_themeManager->setTheme(static_cast<ThemeManager::Theme>(themeIdx));
    else
        m_themeManager->setTheme(ThemeManager::Office2010Blue);

    QByteArray state = m_settings->mainWindowState();
    if (!state.isEmpty())
        restoreState(state);

    rebuildRecentFilesMenu();
}

void MainWindow::saveSettings()
{
    m_settings->setWindowGeometry(geometry());
    m_settings->setWindowMaximized(isMaximized());
    m_settings->setMainWindowState(saveState());
    m_settings->save();
}

void MainWindow::updateWindowTitle()
{
    setWindowTitle(QStringLiteral("Document.Editor"));
}

void MainWindow::closeCurrentDocument()
{
    m_docManager->closeDocument();
    updateEditorActions();
}

void MainWindow::closeAllDocuments()
{
    m_docManager->closeAllDocuments();
    updateEditorActions();
}

void MainWindow::closeAllButCurrent()
{
    m_docManager->closeAllButCurrent();
    updateEditorActions();
}

void MainWindow::closeEvent(QCloseEvent* event)
{
    m_docManager->closeAllDocuments();
    if (m_docManager->tabCount() > 0) {
        event->ignore();
        return;
    }
    saveSettings();
    event->accept();
}

void MainWindow::keyPressEvent(QKeyEvent* event)
{
    if (event->key() == Qt::Key_Insert && (event->modifiers() & Qt::ControlModifier)) {
        event->ignore();
        return;
    }
    QMainWindow::keyPressEvent(event);
}
