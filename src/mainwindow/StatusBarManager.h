#pragma once

#include <QObject>

class QStatusBar;
class QLabel;

class StatusBarManager : public QObject
{
    Q_OBJECT

public:
    explicit StatusBarManager(QStatusBar *statusBar, QObject *parent = nullptr);

    void setLineColumn(int line, int column, int totalLines, int totalColumns);
    void setWordCount(int count);
    void setFileSize(const QString &size);
    void setZoomLevel(int percent);

private:
    QStatusBar *m_statusBar = nullptr;
    QLabel *m_lineColLabel = nullptr;
    QLabel *m_wordCountLabel = nullptr;
    QLabel *m_fileSizeLabel = nullptr;
    QLabel *m_zoomLabel = nullptr;
};
