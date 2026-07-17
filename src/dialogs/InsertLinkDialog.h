#pragma once

#include <QDialog>

class QLineEdit;

class InsertLinkDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertLinkDialog(QWidget *parent = nullptr);

    QString url() const;
    QString displayText() const;

protected:
    void accept() override;

private:
    QLineEdit *m_urlEdit = nullptr;
    QLineEdit *m_textEdit = nullptr;
};
