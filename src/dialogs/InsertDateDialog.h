#pragma once

#include <QDialog>

class QListWidget;

class InsertDateDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertDateDialog(QWidget *parent = nullptr);

    QString dateFormat() const;

private:
    QListWidget *m_formatList = nullptr;
};
