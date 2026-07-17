#pragma once

#include <QDialog>

class QListWidget;

class InsertTimeDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertTimeDialog(QWidget *parent = nullptr);

    QString timeFormat() const;

private:
    QListWidget *m_formatList = nullptr;
};
