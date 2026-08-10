#include "MainWindow.h"
#include "ActionManager.h"
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
#include <QPointer>
#include <QCloseEvent>
#include <QColorDialog>
#include <QComboBox>
#include <QDesktopServices>
#include <QFontDialog>
#include <QImage>
#include <QInputDialog>
#include <QKeyEvent>
#include <QLabel>
#include <QMenuBar>
#include <QMessageBox>
#include <QSlider>
#include <QStatusBar>
#include <QTimer>
#include <QToolBar>
#include <QUrl>

#include "dialogs/AboutDialog.h"
#include "dialogs/FindDialog.h"
#include "dialogs/GoToLineDialog.h"
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

static QTextDocument::FindFlags buildFindFlags(bool matchCase, bool wholeWord = false,
                                               bool backward = false)
{
    QTextDocument::FindFlags flags;
    if (matchCase)
        flags |= QTextDocument::FindCaseSensitively;
    if (wholeWord)
        flags |= QTextDocument::FindWholeWords;
    if (backward)
        flags |= QTextDocument::FindBackward;
    return flags;
}

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
{
    auto *app = qobject_cast<Application*>(QApplication::instance());
    m_settings = app ? app->settings() : new Settings(this);

    m_themeManager = new ThemeManager(this);

    setWindowTitle(QStringLiteral("Document.Editor"));
    setMinimumSize(800, 600);
    resize(1200, 800);

    createCentralArea();

    m_actionManager = new ActionManager(this);
    m_actionManager->setupActions(this);
    m_actionManager->setupToolBars(this);

    createStatusBar();

    m_docManager = new DocumentManager(m_mdiArea, this, this);

    m_docService = new DocumentService(m_docManager, this);

    connect(m_docService, &DocumentService::documentOpened, this, [this](const QString &path) {
        if (!path.isEmpty())
            m_settings->addRecentFile(path);
        connectEditorActionSignals(currentEditor());
        m_actionManager->updateEditorActions(currentEditor(), m_docManager, m_statusBarManager);
        applyEditorSettings(currentEditor());
        applyRulerToTabs();
    });
    connect(m_docService, &DocumentService::statusMessage, this, [this](const QString &msg, int timeout) {
        statusBar()->showMessage(msg, timeout);
    });

    m_actionManager->showRulerAction()->setChecked(m_settings->showRuler());

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
        QPointer<NetworkUtils> safeNet(netUtils);
        connect(netUtils, &NetworkUtils::updateAvailable, this, [this, safeNet](const QString& version, const QString& url) {
            if (safeNet)
                safeNet->deleteLater();
            statusBar()->showMessage(
                tr("Version %1 is available. Go to Help > Check for Updates.").arg(version), 8000);
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
    m_spellChecker = std::make_unique<SpellChecker>();

    // Text-to-speech
    m_tts = new QTextToSpeech(this);
    applyTtsSettings();

    // Zoom debounce timer
    m_zoomDebounceTimer = new QTimer(this);
    m_zoomDebounceTimer->setSingleShot(true);
    m_zoomDebounceTimer->setInterval(50);
    connect(m_zoomDebounceTimer, &QTimer::timeout, this, [this]() {
        int pct = qRound(m_zoomPendingLevel * 100.0);
        if (auto *e = m_zoomPendingEditor.data())
            e->setZoomLevel(m_zoomPendingLevel);
        m_actionManager->zoomSlider()->setValue(pct);
        m_actionManager->zoomLabel()->setText(QStringLiteral(" %1%").arg(pct));
        m_statusBarManager->setZoomLevel(pct);
        applyRulerToTabs();
    });

    // Startup behavior: create/open per settings (unless files were passed on the CLI)
    const bool hasStartupFiles = app && !app->startupFiles().isEmpty();
    if (!hasStartupFiles) {
        if (m_settings->startupMode() == 0) {
            QTimer::singleShot(0, this, [this]() { m_docService->newDocument(); });
        } else if (m_settings->startupMode() == 1) {
            QTimer::singleShot(0, this, [this]() { m_docService->openDocument(); });
        }
    }
}

MainWindow::~MainWindow()
{
    // Destroy editors (and their highlighters) before the spell checker they reference.
    if (m_mdiArea)
        qDeleteAll(m_mdiArea->subWindowList());
    if (m_pluginManager)
        m_pluginManager->unloadPlugins();
}

ActionManager* MainWindow::actionManager() const
{
    return m_actionManager;
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

void MainWindow::openFile(const QString& filePath)
{
    if (!filePath.isEmpty())
        m_docService->openDocument(filePath);
}

void MainWindow::applyRulerToTabs()
{
    const RulerWidget::Unit unit = m_settings->rulerMeasurement() == 1
        ? RulerWidget::Centimeters
        : RulerWidget::Inches;
    const bool visible = m_actionManager->showRulerAction()->isChecked();
    for (DocumentTab* tab : m_docManager->allTabs()) {
        tab->setRulerVisible(visible);
        tab->setRulerUnit(unit);
        tab->syncRuler();
    }
}

void MainWindow::applyEditorSettings(DocumentEditor* editor)
{
    if (!editor)
        return;
    editor->setSpellChecker(m_spellChecker.get());
    editor->setSpellCheckEnabled(m_settings->spellCheckEnabled());
    if (editor->document()->toPlainText().isEmpty()) {
        QFont f = m_settings->defaultFont();
        if (f.family().isEmpty())
            f = QFont(QStringLiteral("Segoe UI"));
        if (f.pointSizeF() <= 0)
            f.setPointSizeF(m_settings->defaultFontSize());
        editor->setBaseFont(f);
    }
}

void MainWindow::applyTtsSettings()
{
    if (!m_tts)
        return;
    const QList<QVoice> voices = m_tts->availableVoices();
    const int idx = m_settings->ttsVoice() - 1;
    if (idx >= 0 && idx < voices.size())
        m_tts->setVoice(voices.at(idx));
    m_tts->setRate(m_settings->ttsSpeed() / 10.0);
}

void MainWindow::applyTheme(ThemeManager::Theme theme)
{
    m_themeManager->setTheme(theme);
    m_settings->setTheme(static_cast<int>(theme));
    const QList<QPair<ThemeManager::Theme, QAction*>> themeActions = {
        { ThemeManager::Office2010Blue,   m_actionManager->themeOffice2010Action() },
        { ThemeManager::Office2010Silver, m_actionManager->themeOffice2010SilverAction() },
        { ThemeManager::Office2010Black,  m_actionManager->themeOffice2010BlackAction() },
        { ThemeManager::Office2013,       m_actionManager->themeOffice2013Action() },
        { ThemeManager::Windows8,         m_actionManager->themeWindows8Action() },
    };
    for (const auto& pair : themeActions)
        pair.second->setChecked(pair.first == theme);
}

void MainWindow::createCentralArea()
{
    m_mdiArea = new QMdiArea(this);
    m_mdiArea->setViewMode(QMdiArea::TabbedView);
    m_mdiArea->setTabsMovable(true);
    m_mdiArea->setDocumentMode(true);

    applyTabSettings();

    connect(m_mdiArea, &QMdiArea::subWindowActivated, this, [this]() {
        m_actionManager->updateEditorActions(currentEditor(), m_docManager, m_statusBarManager);
    });
    setCentralWidget(m_mdiArea);
}

void MainWindow::applyTabSettings()
{
    m_mdiArea->setTabPosition(static_cast<QTabWidget::TabPosition>(m_settings->tabPlacement()));
    m_mdiArea->setTabShape(m_settings->tabSizeMode() == 0
            ? QTabWidget::Rounded
            : QTabWidget::Triangular);
    // Close-button mode: 0 = all tabs, 1 = active tab, 2 = none.
    // QMdiArea has no per-tab close API, so "all" and "active" both enable them.
    m_mdiArea->setTabsClosable(m_settings->tabCloseButtonMode() != 2);
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
    connect(m_actionManager->newAction(), &QAction::triggered, m_docService, &DocumentService::newDocument);
    connect(m_actionManager->openAction(), &QAction::triggered, this, [this]() { m_docService->openDocument(); });
    connect(m_actionManager->closeAction(), &QAction::triggered, this, &MainWindow::closeCurrentDocument);
    connect(m_actionManager->closeAllAction(), &QAction::triggered, this, &MainWindow::closeAllDocuments);
    connect(m_actionManager->closeAllButThisAction(), &QAction::triggered, this, &MainWindow::closeAllButCurrent);
    connect(m_actionManager->saveAction(), &QAction::triggered, this, [this]() { m_docManager->saveDocument(); });
    connect(m_actionManager->saveAsAction(), &QAction::triggered, this, [this]() { m_docManager->saveDocumentAs(); });
    connect(m_actionManager->saveCopyAction(), &QAction::triggered, this, [this]() { m_docManager->saveDocumentCopy(); });
    connect(m_actionManager->saveAllAction(), &QAction::triggered, this, [this]() { m_docManager->saveAllDocuments(); });
    connect(m_actionManager->exitAction(), &QAction::triggered, this, &QWidget::close);

    connect(m_actionManager->importFtpAction(), &QAction::triggered, m_docService, &DocumentService::importFtp);
    connect(m_actionManager->importArchiveAction(), &QAction::triggered, m_docService, &DocumentService::importArchive);
    connect(m_actionManager->importImageAction(), &QAction::triggered, m_docService, &DocumentService::importImage);

    connect(m_actionManager->exportWordpressAction(), &QAction::triggered, m_docService, &DocumentService::exportWordpress);
    connect(m_actionManager->exportEmailAction(), &QAction::triggered, m_docService, &DocumentService::exportEmail);
    connect(m_actionManager->exportFtpAction(), &QAction::triggered, m_docService, &DocumentService::exportFtp);
    connect(m_actionManager->exportPdfAction(), &QAction::triggered, m_docService, &DocumentService::exportPdf);
    connect(m_actionManager->exportArchiveAction(), &QAction::triggered, m_docService, &DocumentService::exportArchive);
    connect(m_actionManager->exportImageAction(), &QAction::triggered, m_docService, &DocumentService::exportImage);
    connect(m_actionManager->exportSoundAction(), &QAction::triggered, m_docService, &DocumentService::exportSound);

    connect(m_actionManager->printAction(), &QAction::triggered, m_docService, &DocumentService::printDocument);
    connect(m_actionManager->pageSetupAction(), &QAction::triggered, m_docService, &DocumentService::pageSetup);

    connect(m_actionManager->revertAction(), &QAction::triggered, m_docService, &DocumentService::revertDocument);

    connect(m_actionManager->optionsAction(), &QAction::triggered, this, [this]() {
        OptionsDialog dlg(m_settings, this);
        if (dlg.exec() == QDialog::Accepted) {
            m_settings->save();
            int themeIdx = m_settings->theme();
            if (themeIdx >= ThemeManager::Office2010Blue && themeIdx <= ThemeManager::Windows8)
                applyTheme(static_cast<ThemeManager::Theme>(themeIdx));
            applyTabSettings();
            applyRulerToTabs();
            applyEditorSettings(currentEditor());
            applyTtsSettings();
            rebuildRecentFilesMenu();
            m_actionManager->updateEditorActions(currentEditor(), m_docManager, m_statusBarManager);
        }
    });
    connect(m_settings, &Settings::recentFilesChanged, this, &MainWindow::rebuildRecentFilesMenu);
}

void MainWindow::connectEditActions()
{
    connect(m_actionManager->undoAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->undo(); });
    connect(m_actionManager->redoAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->redo(); });
    connect(m_actionManager->cutAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->cut(); });
    connect(m_actionManager->copyAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->copy(); });
    connect(m_actionManager->pasteAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->paste(); });
    connect(m_actionManager->pasteTextAction(), &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor()) {
            QString text = QApplication::clipboard()->text();
            if (!text.isEmpty())
                e->textCursor().insertText(text);
        }
    });
    connect(m_actionManager->pasteImageAction(), &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor()) {
            QImage img = QApplication::clipboard()->image();
            if (!img.isNull())
                e->textCursor().insertImage(img);
        }
    });
    connect(m_actionManager->deleteAction(), &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor()) {
            QTextCursor c = e->textCursor();
            c.removeSelectedText();
        }
    });
    connect(m_actionManager->selectAllAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->selectAll(); });
    connect(m_actionManager->findAction(), &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        FindDialog dlg(this);
        connect(&dlg, &FindDialog::findNext, this, [this](const QString& text, bool matchCase, bool wholeWord) {
            currentEditor()->find(text, buildFindFlags(matchCase, wholeWord));
        });
        connect(&dlg, &FindDialog::findPrevious, this, [this](const QString& text, bool matchCase, bool wholeWord) {
            currentEditor()->find(text, buildFindFlags(matchCase, wholeWord, true));
        });
        dlg.exec();
    });
    connect(m_actionManager->replaceAction(), &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        ReplaceDialog dlg(this);
        connect(&dlg, &ReplaceDialog::findNext, this, [this](const QString& text, bool matchCase) {
            currentEditor()->find(text, buildFindFlags(matchCase));
        });
        connect(&dlg, &ReplaceDialog::replace, this, [this](const QString& find, const QString& replace, bool matchCase) {
            if (currentEditor()->find(find, buildFindFlags(matchCase)))
                currentEditor()->textCursor().insertText(replace);
        });
        connect(&dlg, &ReplaceDialog::replaceAll, this, [this](const QString& find, const QString& replace, bool matchCase) {
            if (find.isEmpty())
                return;
            const QTextDocument::FindFlags flags = buildFindFlags(matchCase);
            int count = 0;
            QTextCursor cursor = currentEditor()->textCursor();
            cursor.movePosition(QTextCursor::Start);
            while (!cursor.isNull() && cursor.position() < currentEditor()->document()->characterCount()) {
                cursor = currentEditor()->document()->find(find, cursor, flags);
                if (cursor.isNull())
                    break;
                int prevPos = cursor.position();
                cursor.insertText(replace);
                ++count;
                if (cursor.position() == prevPos)
                    break;
                if (count > 100000)
                    break;
            }
        });
        dlg.exec();
    });
    connect(m_actionManager->goToAction(), &QAction::triggered, this, [this]() {
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
    connect(m_actionManager->boldAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleBold(); });
    connect(m_actionManager->italicAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleItalic(); });
    connect(m_actionManager->underlineAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleUnderline(); });
    connect(m_actionManager->strikethroughAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleStrikethrough(); });
    connect(m_actionManager->subscriptAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleSubscript(); });
    connect(m_actionManager->superscriptAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleSuperscript(); });
    connect(m_actionManager->clearFormattingAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->clearFormatting(); });
    connect(m_actionManager->fontFaceAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) {
            bool ok = false;
            QFont font = QFontDialog::getFont(&ok, e->currentFont(), this);
            if (ok) {
                QTextCharFormat fmt;
                fmt.setFont(font);
                e->textCursor().mergeCharFormat(fmt);
            }
        }
    });
    connect(m_actionManager->fontSizeAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) {
            bool ok = false;
            int size = QInputDialog::getInt(this, tr("Font Size"), tr("Size:"),
                qRound(e->currentFont().pointSizeF()), 1, 999, 1, &ok);
            if (ok) {
                QTextCharFormat fmt;
                fmt.setFontPointSize(size);
                e->textCursor().mergeCharFormat(fmt);
            }
        }
    });
    connect(m_actionManager->fontColorAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) {
            QColor color = QColorDialog::getColor(e->textColor(), this, tr("Font Color"));
            if (color.isValid()) {
                QTextCharFormat fmt;
                fmt.setForeground(color);
                e->textCursor().mergeCharFormat(fmt);
            }
        }
    });
    connect(m_actionManager->highlightColorAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) {
            QColor color = QColorDialog::getColor(e->textBackgroundColor(),
                this, tr("Highlight Color"));
            if (color.isValid()) {
                QTextCharFormat fmt;
                fmt.setBackground(color);
                e->textCursor().mergeCharFormat(fmt);
            }
        }
    });
    connect(m_actionManager->alignLeftAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignLeft); });
    connect(m_actionManager->alignCenterAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignCenter); });
    connect(m_actionManager->alignRightAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignRight); });
    connect(m_actionManager->alignJustifyAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->setParagraphAlignment(Qt::AlignJustify); });
    connect(m_actionManager->bulletListAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleBulletList(); });
    connect(m_actionManager->numberListAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->toggleNumberList(); });
    connect(m_actionManager->indentMoreAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->indentMore(); });
    connect(m_actionManager->indentLessAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor()) e->indentLess(); });
    connect(m_actionManager->lineSpacingAction(), &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        qreal current = currentEditor()->textCursor().blockFormat().lineHeight() / 100.0;
        LineSpacingDialog dlg(current, this);
        if (dlg.exec() == QDialog::Accepted)
            currentEditor()->setLineSpacing(dlg.lineSpacing());
    });
    connect(m_actionManager->leftToRightAction(), &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor())
            e->setTextDirection(Qt::LeftToRight);
    });
    connect(m_actionManager->rightToLeftAction(), &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor())
            e->setTextDirection(Qt::RightToLeft);
    });

    connect(m_actionManager->fontCombo(), &QComboBox::currentTextChanged, this, [this](const QString& family) {
        if (m_actionManager->isUpdatingFont() || !currentEditor())
            return;
        QTextCharFormat fmt;
        fmt.setFontFamilies({ family });
        currentEditor()->textCursor().mergeCharFormat(fmt);
    });
    connect(m_actionManager->fontSizeCombo(), &QComboBox::currentTextChanged, this, [this](const QString& text) {
        if (m_actionManager->isUpdatingFont() || !currentEditor())
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
    connect(m_actionManager->insertTableAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertTableInteractive(this, e);
    });
    connect(m_actionManager->insertImageAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertImageInteractive(this, e);
    });
    connect(m_actionManager->insertShapeAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertShapeInteractive(this, e);
    });
    connect(m_actionManager->insertChartAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertChartInteractive(this, e);
    });
    connect(m_actionManager->insertLinkAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertLinkInteractive(this, e);
    });
    connect(m_actionManager->insertSymbolAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertSymbolInteractive(this, e);
    });
    connect(m_actionManager->insertHorizontalLineAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertHorizontalLineInteractive(e);
    });
    connect(m_actionManager->insertDateAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertDateInteractive(this, e);
    });
    connect(m_actionManager->insertTimeAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertTimeInteractive(this, e);
    });
    connect(m_actionManager->insertVideoAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertVideoInteractive(this, e);
    });
    connect(m_actionManager->insertHeaderAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertHeader(e);
    });
    connect(m_actionManager->insertFooterAction(), &QAction::triggered, this, [this]() {
        if (auto *e = currentEditor())
            DocumentService::insertFooter(e);
    });
}

void MainWindow::connectPageLayoutActions()
{
    connect(m_actionManager->pageSizeAction(), &QAction::triggered, this, [this]() {
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
            bool wOk = false, hOk = false;
            qreal w = parts[0].trimmed().toDouble(&wOk);
            qreal h = parts[1].trimmed().toDouble(&hOk);
            if (wOk && hOk && w > 0 && h > 0) {
                currentEditor()->setPageWidth(w);
                currentEditor()->setPageHeight(h);
            } else {
                QMessageBox::warning(this, tr("Page Size"),
                    tr("Invalid page size. Enter positive numbers."));
            }
        }
    });
    connect(m_actionManager->pageMarginsAction(), &QAction::triggered, this, [this]() {
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
            bool ok = false;
            const qreal left = parts[0].trimmed().toDouble(&ok);
            if (!ok || left < 0) {
                QMessageBox::warning(this, tr("Page Margins"),
                    tr("Invalid margins. Enter non-negative numbers."));
                return;
            }
            qreal top = 0, right = 0, bottom = 0;
            top = parts[1].trimmed().toDouble(&ok);
            if (!ok || top < 0) {
                QMessageBox::warning(this, tr("Page Margins"),
                    tr("Invalid margins. Enter non-negative numbers."));
                return;
            }
            right = parts[2].trimmed().toDouble(&ok);
            if (!ok || right < 0) {
                QMessageBox::warning(this, tr("Page Margins"),
                    tr("Invalid margins. Enter non-negative numbers."));
                return;
            }
            bottom = parts[3].trimmed().toDouble(&ok);
            if (!ok || bottom < 0) {
                QMessageBox::warning(this, tr("Page Margins"),
                    tr("Invalid margins. Enter non-negative numbers."));
                return;
            }
            currentEditor()->setPageMargins(QMarginsF(left, top, right, bottom));
        }
    });
    connect(m_actionManager->pageOrientationAction(), &QAction::triggered, this, [this]() {
        if (auto* e = currentEditor())
            e->togglePageOrientation();
    });
    connect(m_actionManager->pageBackgroundAction(), &QAction::triggered, this, [this]() {
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
    connect(m_actionManager->spellCheckAction(), &QAction::triggered, this, [this]() {
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
            SpellCheckDialog dlg(m_spellChecker.get(), misspelled, this);
            if (dlg.exec() == QDialog::Accepted) {
                const auto reps = dlg.replacements();
                for (const auto& rep : reps) {
                    if (rep.original.isEmpty())
                        continue;
                    QTextCursor cursor(e->document());
                    cursor.movePosition(QTextCursor::Start);
                    cursor = e->document()->find(rep.original, cursor);
                    int replaced = 0;
                    while (!cursor.isNull() && replaced < 1000) {
                        cursor.insertText(rep.replacement);
                        ++replaced;
                        int pos = cursor.position();
                        cursor = e->document()->find(rep.original, cursor);
                        if (!cursor.isNull() && cursor.position() == pos)
                            break;
                    }
                }
            }
        }
    });
    connect(m_actionManager->translateAction(), &QAction::triggered, this, [this]() {
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
    connect(m_actionManager->wordCountAction(), &QAction::triggered, this, [this]() {
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
    connect(m_actionManager->ttsAction(), &QAction::triggered, this, [this]() {
        if (!currentEditor())
            return;
        if (m_tts->state() == QTextToSpeech::Speaking) {
            m_tts->stop();
            m_actionManager->ttsAction()->setText(tr("Speak"));
            return;
        }
        QString text = currentEditor()->textCursor().selectedText();
        if (text.isEmpty())
            text = currentEditor()->toPlainText();
        if (!text.isEmpty()) {
            m_tts->say(text);
            m_actionManager->ttsAction()->setText(tr("Stop"));
        }
    });
    connect(m_tts, &QTextToSpeech::stateChanged, this, [this](QTextToSpeech::State state) {
        if (state == QTextToSpeech::Ready || state == QTextToSpeech::Error)
            m_actionManager->ttsAction()->setText(tr("Speak"));
    });
}

void MainWindow::connectViewActions()
{
    connect(m_actionManager->showRulerAction(), &QAction::toggled, this, [this](bool checked) {
        m_settings->setShowRuler(checked);
        applyRulerToTabs();
    });
    connect(m_actionManager->showStatusBarAction(), &QAction::toggled, this, [this](bool checked) {
        statusBar()->setVisible(checked);
        m_settings->setShowStatusBar(checked);
    });
    connect(m_actionManager->themeOffice2010Action(), &QAction::triggered, this, [this]() {
        applyTheme(ThemeManager::Office2010Blue);
    });
    connect(m_actionManager->themeOffice2010SilverAction(), &QAction::triggered, this, [this]() {
        applyTheme(ThemeManager::Office2010Silver);
    });
    connect(m_actionManager->themeOffice2010BlackAction(), &QAction::triggered, this, [this]() {
        applyTheme(ThemeManager::Office2010Black);
    });
    connect(m_actionManager->themeOffice2013Action(), &QAction::triggered, this, [this]() {
        applyTheme(ThemeManager::Office2013);
    });
    connect(m_actionManager->themeWindows8Action(), &QAction::triggered, this, [this]() {
        applyTheme(ThemeManager::Windows8);
    });

    // Zoom connections
    auto updateZoomDisplay = [this](qreal level) {
        int pct = qRound(level * 100.0);
        m_actionManager->zoomSlider()->setValue(pct);
        m_actionManager->zoomLabel()->setText(QStringLiteral(" %1%").arg(pct));
        m_statusBarManager->setZoomLevel(pct);
        applyRulerToTabs();
    };

    connect(m_actionManager->zoomInAction(), &QAction::triggered, this, [this, updateZoomDisplay]() {
        if (auto *e = currentEditor()) {
            qreal level = qMin(e->zoomLevel() + 0.1, 5.0);
            e->setZoomLevel(level);
            updateZoomDisplay(level);
        }
    });

    connect(m_actionManager->zoomOutAction(), &QAction::triggered, this, [this, updateZoomDisplay]() {
        if (auto *e = currentEditor()) {
            qreal level = qMax(e->zoomLevel() - 0.1, 0.1);
            e->setZoomLevel(level);
            updateZoomDisplay(level);
        }
    });

    connect(m_actionManager->zoomSlider(), &QSlider::valueChanged, this, [this](int value) {
        DocumentEditor *e = currentEditor();
        if (!e)
            return;
        m_zoomPendingLevel = value / 100.0;
        m_zoomPendingEditor = e;
        m_zoomDebounceTimer->stop();
        m_zoomDebounceTimer->start();
    });
}

void MainWindow::connectHelpActions()
{
    connect(m_actionManager->aboutAction(), &QAction::triggered, this, [this]() {
        AboutDialog dlg(this);
        dlg.exec();
    });
    connect(m_actionManager->onlineHelpAction(), &QAction::triggered, this, []() {
        QDesktopServices::openUrl(QUrl(QStringLiteral("http://documenteditor.net/documentation/")));
    });
    connect(m_actionManager->websiteAction(), &QAction::triggered, this, []() {
        QDesktopServices::openUrl(QUrl(QStringLiteral("http://documenteditor.net")));
    });
    connect(m_actionManager->checkUpdatesAction(), &QAction::triggered, this, [this]() {
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
    connect(m_actionManager->pluginsAction(), &QAction::triggered, this, [this]() {
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

void MainWindow::connectEditorActionSignals(DocumentEditor* editor)
{
    if (m_actionSignalsEditor == editor)
        return;
    if (m_actionSignalsEditor) {
        disconnect(m_actionSignalsEditor, &QTextEdit::textChanged, this, nullptr);
        disconnect(m_actionSignalsEditor, &QTextEdit::cursorPositionChanged, this, nullptr);
    }
    m_actionSignalsEditor = editor;
    if (!editor)
        return;
    connect(editor, &QTextEdit::textChanged, this, [this]() {
        if (auto *e = currentEditor())
            m_actionManager->updateEditorActions(e, m_docManager, m_statusBarManager);
    });
    connect(editor, &QTextEdit::cursorPositionChanged, this, [this]() {
        if (auto *e = currentEditor())
            m_actionManager->updateEditorActions(e, m_docManager, m_statusBarManager);
    });
}

void MainWindow::connectDocumentSignals()
{
    connect(m_docManager, &DocumentManager::activeTabChanged, this, [this](DocumentTab* tab) {
        connectEditorActionSignals(tab ? tab->editor() : nullptr);
        m_actionManager->updateEditorActions(currentEditor(), m_docManager, m_statusBarManager);
        if (tab) {
            applyEditorSettings(tab->editor());
            applyRulerToTabs();
        }
    });
    connect(m_docManager, &DocumentManager::statusLineColumnChanged,
        m_statusBarManager, &StatusBarManager::setLineColumn);
    connect(m_docManager, &DocumentManager::statusWordCountChanged,
        m_statusBarManager, &StatusBarManager::setWordCount);
}

void MainWindow::rebuildRecentFilesMenu()
{
    m_actionManager->recentFilesMenu()->clear();
    if (!m_settings->showRecentDocuments())
        return;
    QStringList recent = m_settings->recentFiles();
    if (recent.isEmpty()) {
        QAction* emptyAction = m_actionManager->recentFilesMenu()->addAction(tr("No recent files"));
        emptyAction->setEnabled(false);
        return;
    }
    for (const QString& file : recent) {
        QAction* action = m_actionManager->recentFilesMenu()->addAction(file);
        connect(action, &QAction::triggered, this, [this, file]() {
            m_docService->openDocument(file);
        });
    }
    m_actionManager->recentFilesMenu()->addSeparator();
    QAction* clearAction = m_actionManager->recentFilesMenu()->addAction(tr("Clear Recent Files"));
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

    statusBar()->setVisible(m_settings->showStatusBar());
    m_actionManager->showStatusBarAction()->setChecked(m_settings->showStatusBar());
    applyTabSettings();

    if (m_settings->enableGlass())
        setAttribute(Qt::WA_TranslucentBackground);

    int themeIdx = m_settings->theme();
    if (themeIdx >= ThemeManager::Office2010Blue && themeIdx <= ThemeManager::Windows8)
        applyTheme(static_cast<ThemeManager::Theme>(themeIdx));
    else
        applyTheme(ThemeManager::Office2010Blue);

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

void MainWindow::closeCurrentDocument()
{
    m_docManager->closeDocument();
    m_actionManager->updateEditorActions(currentEditor(), m_docManager, m_statusBarManager);
}

void MainWindow::closeAllDocuments()
{
    m_docManager->closeAllDocuments();
    m_actionManager->updateEditorActions(currentEditor(), m_docManager, m_statusBarManager);
}

void MainWindow::closeAllButCurrent()
{
    m_docManager->closeAllButCurrent();
    m_actionManager->updateEditorActions(currentEditor(), m_docManager, m_statusBarManager);
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
