#pragma once

#include <QDialog>

class QLineEdit;
class QCheckBox;

class ReplaceDialog : public QDialog
{
    Q_OBJECT

public:
    explicit ReplaceDialog(QWidget *parent = nullptr);

    QString findText() const;
    QString replaceText() const;
    bool matchCase() const;

signals:
    void findNext(const QString &text, bool matchCase);
    void replace(const QString &find, const QString &replace, bool matchCase);
    void replaceAll(const QString &find, const QString &replace, bool matchCase);

private:
    QLineEdit *m_findEdit = nullptr;
    QLineEdit *m_replaceEdit = nullptr;
    QCheckBox *m_matchCaseCheck = nullptr;
};
