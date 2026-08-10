#include "ActionManager.h"
#include "editor/DocumentEditor.h"
#include "editor/DocumentManager.h"
#include "mainwindow/StatusBarManager.h"

#include <QComboBox>
#include <QFileInfo>
#include <QFontDatabase>
#include <QIcon>
#include <QLabel>
#include <QMainWindow>
#include <QMenu>
#include <QMenuBar>
#include <QSlider>
#include <QStyle>
#include <QToolBar>

static QIcon icon(const QString& name)
{
    return QIcon(QStringLiteral(":/icons/%1.svg").arg(name));
}

ActionManager::ActionManager(QObject* parent)
    : QObject(parent)
{
}

// ============================================================
// Getter implementations
// ============================================================

// File actions
QAction* ActionManager::newAction() const { return m_newAction; }
QAction* ActionManager::openAction() const { return m_openAction; }
QAction* ActionManager::closeAction() const { return m_closeAction; }
QAction* ActionManager::closeAllAction() const { return m_closeAllAction; }
QAction* ActionManager::closeAllButThisAction() const { return m_closeAllButThisAction; }
QAction* ActionManager::saveAction() const { return m_saveAction; }
QAction* ActionManager::saveAsAction() const { return m_saveAsAction; }
QAction* ActionManager::saveCopyAction() const { return m_saveCopyAction; }
QAction* ActionManager::saveAllAction() const { return m_saveAllAction; }
QAction* ActionManager::revertAction() const { return m_revertAction; }
QAction* ActionManager::printAction() const { return m_printAction; }
QAction* ActionManager::pageSetupAction() const { return m_pageSetupAction; }
QAction* ActionManager::exitAction() const { return m_exitAction; }

// Import/Export actions
QAction* ActionManager::importFtpAction() const { return m_importFtpAction; }
QAction* ActionManager::importArchiveAction() const { return m_importArchiveAction; }
QAction* ActionManager::importImageAction() const { return m_importImageAction; }
QAction* ActionManager::exportWordpressAction() const { return m_exportWordpressAction; }
QAction* ActionManager::exportEmailAction() const { return m_exportEmailAction; }
QAction* ActionManager::exportFtpAction() const { return m_exportFtpAction; }
QAction* ActionManager::exportPdfAction() const { return m_exportPdfAction; }
QAction* ActionManager::exportArchiveAction() const { return m_exportArchiveAction; }
QAction* ActionManager::exportImageAction() const { return m_exportImageAction; }
QAction* ActionManager::exportSoundAction() const { return m_exportSoundAction; }

// Edit actions
QAction* ActionManager::undoAction() const { return m_undoAction; }
QAction* ActionManager::redoAction() const { return m_redoAction; }
QAction* ActionManager::cutAction() const { return m_cutAction; }
QAction* ActionManager::copyAction() const { return m_copyAction; }
QAction* ActionManager::pasteAction() const { return m_pasteAction; }
QAction* ActionManager::pasteTextAction() const { return m_pasteTextAction; }
QAction* ActionManager::pasteImageAction() const { return m_pasteImageAction; }
QAction* ActionManager::deleteAction() const { return m_deleteAction; }
QAction* ActionManager::selectAllAction() const { return m_selectAllAction; }
QAction* ActionManager::findAction() const { return m_findAction; }
QAction* ActionManager::replaceAction() const { return m_replaceAction; }
QAction* ActionManager::goToAction() const { return m_goToAction; }

// Font actions
QAction* ActionManager::boldAction() const { return m_boldAction; }
QAction* ActionManager::italicAction() const { return m_italicAction; }
QAction* ActionManager::underlineAction() const { return m_underlineAction; }
QAction* ActionManager::strikethroughAction() const { return m_strikethroughAction; }
QAction* ActionManager::subscriptAction() const { return m_subscriptAction; }
QAction* ActionManager::superscriptAction() const { return m_superscriptAction; }
QAction* ActionManager::clearFormattingAction() const { return m_clearFormattingAction; }
QAction* ActionManager::fontFaceAction() const { return m_fontFaceAction; }
QAction* ActionManager::fontSizeAction() const { return m_fontSizeAction; }
QAction* ActionManager::fontColorAction() const { return m_fontColorAction; }
QAction* ActionManager::highlightColorAction() const { return m_highlightColorAction; }
QComboBox* ActionManager::fontCombo() const { return m_fontCombo; }
QComboBox* ActionManager::fontSizeCombo() const { return m_fontSizeCombo; }

// Paragraph actions
QAction* ActionManager::alignLeftAction() const { return m_alignLeftAction; }
QAction* ActionManager::alignCenterAction() const { return m_alignCenterAction; }
QAction* ActionManager::alignRightAction() const { return m_alignRightAction; }
QAction* ActionManager::alignJustifyAction() const { return m_alignJustifyAction; }
QAction* ActionManager::bulletListAction() const { return m_bulletListAction; }
QAction* ActionManager::numberListAction() const { return m_numberListAction; }
QAction* ActionManager::indentMoreAction() const { return m_indentMoreAction; }
QAction* ActionManager::indentLessAction() const { return m_indentLessAction; }
QAction* ActionManager::lineSpacingAction() const { return m_lineSpacingAction; }
QAction* ActionManager::leftToRightAction() const { return m_leftToRightAction; }
QAction* ActionManager::rightToLeftAction() const { return m_rightToLeftAction; }

// Insert actions
QAction* ActionManager::insertTableAction() const { return m_insertTableAction; }
QAction* ActionManager::insertImageAction() const { return m_insertImageAction; }
QAction* ActionManager::insertShapeAction() const { return m_insertShapeAction; }
QAction* ActionManager::insertChartAction() const { return m_insertChartAction; }
QAction* ActionManager::insertLinkAction() const { return m_insertLinkAction; }
QAction* ActionManager::insertSymbolAction() const { return m_insertSymbolAction; }
QAction* ActionManager::insertHorizontalLineAction() const { return m_insertHorizontalLineAction; }
QAction* ActionManager::insertDateAction() const { return m_insertDateAction; }
QAction* ActionManager::insertTimeAction() const { return m_insertTimeAction; }
QAction* ActionManager::insertVideoAction() const { return m_insertVideoAction; }
QAction* ActionManager::insertHeaderAction() const { return m_insertHeaderAction; }
QAction* ActionManager::insertFooterAction() const { return m_insertFooterAction; }

// Page Layout actions
QAction* ActionManager::pageSizeAction() const { return m_pageSizeAction; }
QAction* ActionManager::pageMarginsAction() const { return m_pageMarginsAction; }
QAction* ActionManager::pageOrientationAction() const { return m_pageOrientationAction; }
QAction* ActionManager::pageBackgroundAction() const { return m_pageBackgroundAction; }

// Review actions
QAction* ActionManager::spellCheckAction() const { return m_spellCheckAction; }
QAction* ActionManager::translateAction() const { return m_translateAction; }
QAction* ActionManager::wordCountAction() const { return m_wordCountAction; }
QAction* ActionManager::ttsAction() const { return m_ttsAction; }

// Options action
QAction* ActionManager::optionsAction() const { return m_optionsAction; }

// View actions
QAction* ActionManager::zoomInAction() const { return m_zoomInAction; }
QAction* ActionManager::zoomOutAction() const { return m_zoomOutAction; }
QSlider* ActionManager::zoomSlider() const { return m_zoomSlider; }
QLabel* ActionManager::zoomLabel() const { return m_zoomLabel; }
QAction* ActionManager::showRulerAction() const { return m_showRulerAction; }
QAction* ActionManager::showStatusBarAction() const { return m_showStatusBarAction; }

// Help actions
QAction* ActionManager::aboutAction() const { return m_aboutAction; }
QAction* ActionManager::checkUpdatesAction() const { return m_checkUpdatesAction; }
QAction* ActionManager::onlineHelpAction() const { return m_onlineHelpAction; }
QAction* ActionManager::websiteAction() const { return m_websiteAction; }
QAction* ActionManager::pluginsAction() const { return m_pluginsAction; }

// Theme actions
QAction* ActionManager::themeOffice2010Action() const { return m_themeOffice2010Action; }
QAction* ActionManager::themeOffice2010SilverAction() const { return m_themeOffice2010SilverAction; }
QAction* ActionManager::themeOffice2010BlackAction() const { return m_themeOffice2010BlackAction; }
QAction* ActionManager::themeOffice2013Action() const { return m_themeOffice2013Action; }
QAction* ActionManager::themeWindows8Action() const { return m_themeWindows8Action; }

// Menus
QMenu* ActionManager::fileMenu() const { return m_fileMenu; }
QMenu* ActionManager::editMenu() const { return m_editMenu; }
QMenu* ActionManager::insertMenu() const { return m_insertMenu; }
QMenu* ActionManager::formatMenu() const { return m_formatMenu; }
QMenu* ActionManager::viewMenu() const { return m_viewMenu; }
QMenu* ActionManager::helpMenu() const { return m_helpMenu; }
QMenu* ActionManager::recentFilesMenu() const { return m_recentFilesMenu; }
QMenu* ActionManager::themeMenu() const { return m_themeMenu; }
QMenu* ActionManager::importMenu() const { return m_importMenu; }
QMenu* ActionManager::exportMenu() const { return m_exportMenu; }

// Toolbars
QToolBar* ActionManager::quickAccessToolBar() const { return m_quickAccessToolBar; }
QToolBar* ActionManager::clipboardToolBar() const { return m_clipboardToolBar; }
QToolBar* ActionManager::fontToolBar() const { return m_fontToolBar; }
QToolBar* ActionManager::paragraphToolBar() const { return m_paragraphToolBar; }
QToolBar* ActionManager::editingToolBar() const { return m_editingToolBar; }
QToolBar* ActionManager::insertToolBar() const { return m_insertToolBar; }
QToolBar* ActionManager::pageLayoutToolBar() const { return m_pageLayoutToolBar; }
QToolBar* ActionManager::reviewToolBar() const { return m_reviewToolBar; }
QToolBar* ActionManager::zoomToolBar() const { return m_zoomToolBar; }

bool ActionManager::isUpdatingFont() const { return m_updatingFont; }
void ActionManager::setUpdatingFont(bool updating) { m_updatingFont = updating; }

void ActionManager::setupActions(QMainWindow* mainWindow)
{
    QMenuBar* mb = mainWindow->menuBar();

    // File menu
    m_fileMenu = mb->addMenu(QObject::tr("&File"));

    m_newAction = m_fileMenu->addAction(QObject::tr("&New"));
    m_newAction->setShortcut(QKeySequence::New);
    m_newAction->setIcon(mainWindow->style()->standardIcon(QStyle::SP_FileIcon));

    m_openAction = m_fileMenu->addAction(QObject::tr("&Open..."));
    m_openAction->setShortcut(QKeySequence::Open);
    m_openAction->setIcon(mainWindow->style()->standardIcon(QStyle::SP_DialogOpenButton));

    m_fileMenu->addSeparator();

    m_closeAction = m_fileMenu->addAction(QObject::tr("&Close"));
    m_closeAction->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_W));
    m_closeAllButThisAction = m_fileMenu->addAction(QObject::tr("Close All But This"));
    m_closeAllAction = m_fileMenu->addAction(QObject::tr("Close All"));

    m_fileMenu->addSeparator();

    m_saveAction = m_fileMenu->addAction(QObject::tr("&Save"));
    m_saveAction->setShortcut(QKeySequence::Save);
    m_saveAction->setIcon(mainWindow->style()->standardIcon(QStyle::SP_DialogSaveButton));
    m_saveAsAction = m_fileMenu->addAction(QObject::tr("Save &As..."));
    m_saveAsAction->setShortcut(QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_S));
    m_saveCopyAction = m_fileMenu->addAction(QObject::tr("Save Cop&y..."));
    m_saveAllAction = m_fileMenu->addAction(QObject::tr("Save A&ll"));
    m_revertAction = m_fileMenu->addAction(QObject::tr("Re&vert"));

    m_fileMenu->addSeparator();

    m_importMenu = m_fileMenu->addMenu(QObject::tr("&Import"));
    m_importFtpAction = m_importMenu->addAction(QObject::tr("From &FTP..."));
    m_importArchiveAction = m_importMenu->addAction(QObject::tr("From &Archive..."));
    m_importImageAction = m_importMenu->addAction(QObject::tr("&Image..."));

    m_exportMenu = m_fileMenu->addMenu(QObject::tr("E&xport"));
    m_exportWordpressAction = m_exportMenu->addAction(QObject::tr("To &Wordpress..."));
    m_exportEmailAction = m_exportMenu->addAction(QObject::tr("To &Email..."));
    m_exportFtpAction = m_exportMenu->addAction(QObject::tr("To &FTP..."));
    m_exportMenu->addSeparator();
    m_exportPdfAction = m_exportMenu->addAction(QObject::tr("To &PDF..."));
    m_exportArchiveAction = m_exportMenu->addAction(QObject::tr("To &Archive..."));
    m_exportImageAction = m_exportMenu->addAction(QObject::tr("To &Image..."));
    m_exportSoundAction = m_exportMenu->addAction(QObject::tr("To &Sound..."));

    m_fileMenu->addSeparator();

    m_recentFilesMenu = m_fileMenu->addMenu(QObject::tr("&Recent Files"));

    m_fileMenu->addSeparator();

    m_optionsAction = m_fileMenu->addAction(QObject::tr("&Options..."));
    m_optionsAction->setMenuRole(QAction::PreferencesRole);

    m_fileMenu->addSeparator();

    m_printAction = m_fileMenu->addAction(QObject::tr("&Print..."));
    m_printAction->setShortcut(QKeySequence::Print);
    m_pageSetupAction = m_fileMenu->addAction(QObject::tr("Page Set&up..."));

    m_fileMenu->addSeparator();

    m_exitAction = m_fileMenu->addAction(QObject::tr("E&xit"));
    m_exitAction->setShortcut(QKeySequence::Quit);

    // Edit menu
    m_editMenu = mb->addMenu(QObject::tr("&Edit"));

    m_undoAction = m_editMenu->addAction(QObject::tr("&Undo"));
    m_undoAction->setShortcut(QKeySequence::Undo);
    m_redoAction = m_editMenu->addAction(QObject::tr("&Redo"));
    m_redoAction->setShortcut(QKeySequence::Redo);

    m_editMenu->addSeparator();

    m_cutAction = m_editMenu->addAction(QObject::tr("Cu&t"));
    m_cutAction->setShortcut(QKeySequence::Cut);
    m_copyAction = m_editMenu->addAction(QObject::tr("&Copy"));
    m_copyAction->setShortcut(QKeySequence::Copy);
    m_pasteAction = m_editMenu->addAction(QObject::tr("&Paste"));
    m_pasteAction->setShortcut(QKeySequence::Paste);
    m_pasteTextAction = m_editMenu->addAction(QObject::tr("Paste &Text"));
    m_pasteImageAction = m_editMenu->addAction(QObject::tr("Paste &Image"));
    m_deleteAction = m_editMenu->addAction(QObject::tr("&Delete"));
    m_deleteAction->setShortcut(QKeySequence::Delete);

    m_editMenu->addSeparator();

    m_selectAllAction = m_editMenu->addAction(QObject::tr("Select &All"));
    m_selectAllAction->setShortcut(QKeySequence::SelectAll);

    m_editMenu->addSeparator();

    m_findAction = m_editMenu->addAction(QObject::tr("&Find..."));
    m_findAction->setShortcut(QKeySequence::Find);
    m_replaceAction = m_editMenu->addAction(QObject::tr("&Replace..."));
    m_replaceAction->setShortcut(QKeySequence::Replace);
    m_goToAction = m_editMenu->addAction(QObject::tr("&Go To..."));
    m_goToAction->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_G));

    // Insert menu
    m_insertMenu = mb->addMenu(QObject::tr("&Insert"));

    m_insertTableAction = m_insertMenu->addAction(QObject::tr("&Table..."));
    m_insertMenu->addSeparator();
    m_insertImageAction = m_insertMenu->addAction(QObject::tr("&Image..."));
    m_insertShapeAction = m_insertMenu->addAction(QObject::tr("S&hape..."));
    m_insertChartAction = m_insertMenu->addAction(QObject::tr("&Chart..."));
    m_insertMenu->addSeparator();
    m_insertLinkAction = m_insertMenu->addAction(QObject::tr("&Link..."));
    m_insertSymbolAction = m_insertMenu->addAction(QObject::tr("S&ymbol..."));
    m_insertHorizontalLineAction = m_insertMenu->addAction(QObject::tr("Horizontal &Line..."));
    m_insertDateAction = m_insertMenu->addAction(QObject::tr("&Date..."));
    m_insertTimeAction = m_insertMenu->addAction(QObject::tr("&Time..."));
    m_insertVideoAction = m_insertMenu->addAction(QObject::tr("&Video..."));
    m_insertHeaderAction = m_insertMenu->addAction(QObject::tr("&Header..."));
    m_insertFooterAction = m_insertMenu->addAction(QObject::tr("&Footer..."));

    // Format menu
    m_formatMenu = mb->addMenu(QObject::tr("F&ormat"));

    m_clearFormattingAction = m_formatMenu->addAction(QObject::tr("&Clear Formatting"));

    QMenu* fontSubMenu = m_formatMenu->addMenu(QObject::tr("&Font"));
    m_fontFaceAction = fontSubMenu->addAction(QObject::tr("Font &Face..."));
    m_fontSizeAction = fontSubMenu->addAction(QObject::tr("Font &Size..."));
    m_fontColorAction = fontSubMenu->addAction(QObject::tr("Font &Color..."));
    m_highlightColorAction = fontSubMenu->addAction(QObject::tr("&Highlight Color..."));

    QMenu* styleSubMenu = m_formatMenu->addMenu(QObject::tr("S&tyle"));
    m_boldAction = styleSubMenu->addAction(QObject::tr("&Bold"));
    m_boldAction->setShortcut(QKeySequence::Bold);
    m_boldAction->setCheckable(true);
    m_italicAction = styleSubMenu->addAction(QObject::tr("&Italic"));
    m_italicAction->setShortcut(QKeySequence::Italic);
    m_italicAction->setCheckable(true);
    m_underlineAction = styleSubMenu->addAction(QObject::tr("&Underline"));
    m_underlineAction->setShortcut(QKeySequence::Underline);
    m_underlineAction->setCheckable(true);
    m_strikethroughAction = styleSubMenu->addAction(QObject::tr("&Strikethrough"));
    m_strikethroughAction->setCheckable(true);

    m_subscriptAction = m_formatMenu->addAction(QObject::tr("&Subscript"));
    m_superscriptAction = m_formatMenu->addAction(QObject::tr("Su&perscript"));
    m_formatMenu->addSeparator();
    m_indentMoreAction = m_formatMenu->addAction(QObject::tr("Indent &More"));
    m_indentLessAction = m_formatMenu->addAction(QObject::tr("Indent &Less"));

    QMenu* listSubMenu = m_formatMenu->addMenu(QObject::tr("&List"));
    m_bulletListAction = listSubMenu->addAction(QObject::tr("&Bullet List"));
    m_numberListAction = listSubMenu->addAction(QObject::tr("&Number List"));

    QMenu* alignSubMenu = m_formatMenu->addMenu(QObject::tr("Ali&gn"));
    m_alignLeftAction = alignSubMenu->addAction(QObject::tr("&Left"));
    m_alignCenterAction = alignSubMenu->addAction(QObject::tr("&Center"));
    m_alignRightAction = alignSubMenu->addAction(QObject::tr("&Right"));
    m_alignJustifyAction = alignSubMenu->addAction(QObject::tr("&Justify"));

    m_lineSpacingAction = m_formatMenu->addAction(QObject::tr("&Line Spacing..."));
    m_formatMenu->addSeparator();
    m_leftToRightAction = m_formatMenu->addAction(QObject::tr("&Left to Right"));
    m_rightToLeftAction = m_formatMenu->addAction(QObject::tr("&Right to Left"));

    QMenu* pageLayoutMenu = m_formatMenu->addMenu(QObject::tr("&Page Layout"));
    m_pageSizeAction = pageLayoutMenu->addAction(QObject::tr("Page &Size..."));
    m_pageMarginsAction = pageLayoutMenu->addAction(QObject::tr("&Margins..."));
    m_pageOrientationAction = pageLayoutMenu->addAction(QObject::tr("&Orientation"));
    m_pageBackgroundAction = pageLayoutMenu->addAction(QObject::tr("&Background..."));

    // View menu
    m_viewMenu = mb->addMenu(QObject::tr("&View"));

    m_zoomInAction = m_viewMenu->addAction(QObject::tr("Zoom &In"));
    m_zoomInAction->setShortcut(QKeySequence::ZoomIn);
    m_zoomOutAction = m_viewMenu->addAction(QObject::tr("Zoom &Out"));
    m_zoomOutAction->setShortcut(QKeySequence::ZoomOut);
    m_viewMenu->addSeparator();
    m_showRulerAction = m_viewMenu->addAction(QObject::tr("Show &Ruler"));
    m_showRulerAction->setCheckable(true);
    m_showRulerAction->setChecked(true);
    m_showStatusBarAction = m_viewMenu->addAction(QObject::tr("Show &Status Bar"));
    m_showStatusBarAction->setCheckable(true);
    m_showStatusBarAction->setChecked(true);
    m_viewMenu->addSeparator();

    m_themeMenu = m_viewMenu->addMenu(QObject::tr("&Theme"));
    m_themeOffice2010Action = m_themeMenu->addAction(QObject::tr("Office 2010 &Blue"));
    m_themeOffice2010Action->setCheckable(true);
    m_themeOffice2010SilverAction = m_themeMenu->addAction(QObject::tr("Office 2010 &Silver"));
    m_themeOffice2010SilverAction->setCheckable(true);
    m_themeOffice2010BlackAction = m_themeMenu->addAction(QObject::tr("Office 2010 &Black"));
    m_themeOffice2010BlackAction->setCheckable(true);
    m_themeOffice2013Action = m_themeMenu->addAction(QObject::tr("Office &2013"));
    m_themeOffice2013Action->setCheckable(true);
    m_themeWindows8Action = m_themeMenu->addAction(QObject::tr("&Windows 8"));
    m_themeWindows8Action->setCheckable(true);

    // Help menu
    m_helpMenu = mb->addMenu(QObject::tr("&Help"));

    m_onlineHelpAction = m_helpMenu->addAction(QObject::tr("&Online Help..."));
    m_websiteAction = m_helpMenu->addAction(QObject::tr("&Website..."));
    m_helpMenu->addSeparator();
    m_checkUpdatesAction = m_helpMenu->addAction(QObject::tr("&Check for Updates..."));
    m_helpMenu->addSeparator();
    m_pluginsAction = m_helpMenu->addAction(QObject::tr("&Plugins..."));
    m_helpMenu->addSeparator();
    m_aboutAction = m_helpMenu->addAction(QObject::tr("&About Document.Editor..."));

    setupIcons();
}

void ActionManager::setupIcons()
{
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
    m_exportPdfAction->setIcon(icon(QStringLiteral("xps")));
    m_exportArchiveAction->setIcon(icon(QStringLiteral("archive")));
    m_exportImageAction->setIcon(icon(QStringLiteral("image")));
    m_exportSoundAction->setIcon(icon(QStringLiteral("sound")));

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

    m_zoomInAction->setIcon(icon(QStringLiteral("zoomin")));
    m_zoomOutAction->setIcon(icon(QStringLiteral("zoomout")));
    m_showRulerAction->setIcon(icon(QStringLiteral("ruler")));
    m_showStatusBarAction->setIcon(icon(QStringLiteral("statusbar")));

    m_onlineHelpAction->setIcon(icon(QStringLiteral("help")));
    m_websiteAction->setIcon(icon(QStringLiteral("website")));
    m_checkUpdatesAction->setIcon(icon(QStringLiteral("updates")));
    m_aboutAction->setIcon(icon(QStringLiteral("about")));
}

void ActionManager::setupToolBars(QMainWindow* mainWindow)
{
    m_quickAccessToolBar = mainWindow->addToolBar(QObject::tr("Quick Access"));
    m_quickAccessToolBar->setObjectName(QStringLiteral("QuickAccessToolBar"));
    m_quickAccessToolBar->setMovable(false);
    m_quickAccessToolBar->addAction(m_newAction);
    m_quickAccessToolBar->addAction(m_openAction);
    m_quickAccessToolBar->addAction(m_saveAction);
    m_quickAccessToolBar->addSeparator();
    m_quickAccessToolBar->addAction(m_undoAction);
    m_quickAccessToolBar->addAction(m_redoAction);

    m_clipboardToolBar = mainWindow->addToolBar(QObject::tr("Clipboard"));
    m_clipboardToolBar->setObjectName(QStringLiteral("ClipboardToolBar"));
    m_clipboardToolBar->addAction(m_cutAction);
    m_clipboardToolBar->addAction(m_copyAction);
    m_clipboardToolBar->addAction(m_pasteAction);

    m_fontToolBar = mainWindow->addToolBar(QObject::tr("Font"));
    m_fontToolBar->setObjectName(QStringLiteral("FontToolBar"));

    m_fontCombo = new QComboBox(m_fontToolBar);
    m_fontCombo->setMinimumWidth(150);
    m_fontCombo->setEditable(false);
    const QFontDatabase fontDb;
    const QStringList families = fontDb.families();
    for (const QString& family : families)
        m_fontCombo->addItem(family);
    m_fontToolBar->addWidget(m_fontCombo);

    m_fontSizeCombo = new QComboBox(m_fontToolBar);
    m_fontSizeCombo->setEditable(true);
    m_fontSizeCombo->setMinimumWidth(60);
    const int sizes[] = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
    for (int size : sizes)
        m_fontSizeCombo->addItem(QString::number(size));
    m_fontToolBar->addWidget(m_fontSizeCombo);

    m_fontToolBar->addSeparator();
    m_fontToolBar->addAction(m_boldAction);
    m_fontToolBar->addAction(m_italicAction);
    m_fontToolBar->addAction(m_underlineAction);
    m_fontToolBar->addAction(m_strikethroughAction);
    m_fontToolBar->addSeparator();
    m_fontToolBar->addAction(m_subscriptAction);
    m_fontToolBar->addAction(m_superscriptAction);

    m_paragraphToolBar = mainWindow->addToolBar(QObject::tr("Paragraph"));
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

    m_editingToolBar = mainWindow->addToolBar(QObject::tr("Editing"));
    m_editingToolBar->setObjectName(QStringLiteral("EditingToolBar"));
    m_editingToolBar->addAction(m_findAction);
    m_editingToolBar->addAction(m_replaceAction);
    m_editingToolBar->addAction(m_goToAction);

    m_insertToolBar = mainWindow->addToolBar(QObject::tr("Insert"));
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

    m_pageLayoutToolBar = mainWindow->addToolBar(QObject::tr("Page Layout"));
    m_pageLayoutToolBar->setObjectName(QStringLiteral("PageLayoutToolBar"));
    m_pageLayoutToolBar->addAction(m_pageSizeAction);
    m_pageLayoutToolBar->addAction(m_pageMarginsAction);
    m_pageLayoutToolBar->addAction(m_pageOrientationAction);
    m_pageLayoutToolBar->addAction(m_pageBackgroundAction);

    m_reviewToolBar = mainWindow->addToolBar(QObject::tr("Review"));
    m_reviewToolBar->setObjectName(QStringLiteral("ReviewToolBar"));
    m_spellCheckAction = m_reviewToolBar->addAction(QObject::tr("Spell Check..."));
    m_spellCheckAction->setIcon(icon(QStringLiteral("spellcheck")));
    m_translateAction = m_reviewToolBar->addAction(QObject::tr("Translate..."));
    m_translateAction->setIcon(icon(QStringLiteral("translate")));
    m_wordCountAction = m_reviewToolBar->addAction(QObject::tr("Word Count"));
    m_wordCountAction->setIcon(icon(QStringLiteral("wordcount")));
    m_reviewToolBar->addSeparator();
    m_ttsAction = m_reviewToolBar->addAction(QObject::tr("Speak"));
    m_ttsAction->setIcon(icon(QStringLiteral("sound")));

    m_zoomToolBar = mainWindow->addToolBar(QObject::tr("Zoom"));
    m_zoomToolBar->setObjectName(QStringLiteral("ZoomToolBar"));
    m_zoomToolBar->addAction(m_zoomOutAction);

    m_zoomSlider = new QSlider(Qt::Horizontal, m_zoomToolBar);
    m_zoomSlider->setRange(10, 500);
    m_zoomSlider->setValue(100);
    m_zoomSlider->setFixedWidth(120);
    m_zoomToolBar->addWidget(m_zoomSlider);

    m_zoomToolBar->addAction(m_zoomInAction);

    m_zoomLabel = new QLabel(QStringLiteral(" 100%"), m_zoomToolBar);
    m_zoomToolBar->addWidget(m_zoomLabel);
}

void ActionManager::updateEditorActions(DocumentEditor* editor, DocumentManager* docManager, StatusBarManager* statusBarManager)
{
    bool hasDoc = editor != nullptr;

    m_saveAction->setEnabled(hasDoc && !editor->isReadOnlyFile());
    m_saveAsAction->setEnabled(hasDoc);
    m_saveCopyAction->setEnabled(hasDoc);
    m_saveAllAction->setEnabled(hasDoc && docManager->tabCount() > 1);
    m_closeAction->setEnabled(hasDoc);
    m_closeAllAction->setEnabled(hasDoc && docManager->tabCount() > 1);
    m_closeAllButThisAction->setEnabled(hasDoc && docManager->tabCount() > 1);
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

        const int pct = qRound(editor->zoomLevel() * 100.0);
        m_zoomSlider->setValue(pct);
        m_zoomLabel->setText(QStringLiteral(" %1%").arg(pct));

        const QString name = editor->documentName();
        if (name.isEmpty()) {
            statusBarManager->setFileSize(QStringLiteral("Unsaved"));
        } else {
            const qint64 bytes = QFileInfo(name).size();
            if (bytes >= 1024)
                statusBarManager->setFileSize(tr("%1 KB").arg(bytes / 1024));
            else
                statusBarManager->setFileSize(tr("%1 B").arg(bytes));
        }
    } else {
        statusBarManager->setLineColumn(0, 0, 0, 0);
        statusBarManager->setWordCount(0);
        statusBarManager->setFileSize(QString());

        m_boldAction->setChecked(false);
        m_italicAction->setChecked(false);
        m_underlineAction->setChecked(false);
        m_strikethroughAction->setChecked(false);

        m_updatingFont = true;
        m_fontCombo->setCurrentIndex(-1);
        m_fontSizeCombo->clearEditText();
        m_updatingFont = false;

        m_zoomSlider->setValue(100);
        m_zoomLabel->setText(QStringLiteral(" 100%"));
    }
}
