#include "StatusBarManager.h"

#include <QStatusBar>
#include <QLabel>

StatusBarManager::StatusBarManager(QStatusBar *statusBar, QObject *parent)
    : QObject(parent)
    , m_statusBar(statusBar)
{
    m_lineColLabel = new QLabel(QStringLiteral(" Line: 1 of 1 | Col: 1 of 1 "));
    m_lineColLabel->setMinimumWidth(200);
    m_statusBar->addWidget(m_lineColLabel);

    m_wordCountLabel = new QLabel(QStringLiteral(" Words: 0 "));
    m_wordCountLabel->setMinimumWidth(100);
    m_statusBar->addWidget(m_wordCountLabel);

    m_fileSizeLabel = new QLabel(QStringLiteral(" Size: 0 KB "));
    m_fileSizeLabel->setMinimumWidth(100);
    m_statusBar->addPermanentWidget(m_fileSizeLabel);

    m_zoomLabel = new QLabel(QStringLiteral(" 100% "));
    m_zoomLabel->setMinimumWidth(60);
    m_statusBar->addPermanentWidget(m_zoomLabel);
}

void StatusBarManager::setLineColumn(int line, int column, int totalLines, int totalColumns)
{
    m_lineColLabel->setText(QStringLiteral(" Line: %1 of %2 | Col: %3 of %4 ")
        .arg(line).arg(totalLines).arg(column).arg(totalColumns));
}

void StatusBarManager::setWordCount(int count)
{
    m_wordCountLabel->setText(QStringLiteral(" Words: %1 ").arg(count));
}

void StatusBarManager::setFileSize(const QString &size)
{
    m_fileSizeLabel->setText(QStringLiteral(" Size: %1 ").arg(size));
}

void StatusBarManager::setZoomLevel(int percent)
{
    m_zoomLabel->setText(QStringLiteral(" %1% ").arg(percent));
}
