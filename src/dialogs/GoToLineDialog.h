#pragma once

#include <QDialog>

class QSpinBox;
class QLabel;

class GoToLineDialog : public QDialog
{
    Q_OBJECT

public:
    explicit GoToLineDialog(int currentLine, int totalLines, QWidget *parent = nullptr);

    int lineNumber() const;

private:
    QSpinBox *m_lineSpin = nullptr;
};
