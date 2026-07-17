#pragma once

#include <QDialog>

class QLineEdit;
class QCheckBox;
class QPushButton;

class FindDialog : public QDialog
{
    Q_OBJECT

public:
    explicit FindDialog(QWidget *parent = nullptr);

    QString searchText() const;
    bool matchCase() const;
    bool wholeWord() const;

signals:
    void findNext(const QString &text, bool matchCase, bool wholeWord);
    void findPrevious(const QString &text, bool matchCase, bool wholeWord);

private:
    QLineEdit *m_searchEdit = nullptr;
    QCheckBox *m_matchCaseCheck = nullptr;
    QCheckBox *m_wholeWordCheck = nullptr;
};
