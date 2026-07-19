#pragma once

#include <QWidget>
#include <QString>

class QScrollArea;
class QVBoxLayout;
class QMenu;

class DocumentEditor;
class RulerWidget;

#include "widgets/RulerWidget.h"

class DocumentTab : public QWidget
{
    Q_OBJECT

public:
    explicit DocumentTab(const QString &title = QString(),
                         QWidget *parent = nullptr);

    DocumentEditor *editor() const;

    QString tabTitle() const;
    void setTabTitle(const QString &title);

    QString documentName() const;
    void setDocumentName(const QString &name);

    void setPageSize(qreal width, qreal height);

    RulerWidget *ruler() const;
    void setRulerVisible(bool visible);
    void setRulerUnit(RulerWidget::Unit unit);

signals:
    void titleChanged(const QString &title);

private:
    void createContextMenu();
    void updateTitle();

    DocumentEditor *m_editor = nullptr;
    QScrollArea *m_scrollArea = nullptr;
    QMenu *m_contextMenu = nullptr;
    RulerWidget *m_ruler = nullptr;
    QString m_tabTitle;
    QString m_documentName;
};
