#pragma once

#include <QDialog>

class QLineEdit;

class InsertVideoDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertVideoDialog(QWidget *parent = nullptr);

    QString videoPath() const;

protected:
    void accept() override;

private:
    QLineEdit *m_pathEdit = nullptr;
    void browseVideo();
};
