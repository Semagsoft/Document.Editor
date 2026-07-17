#pragma once

#include <QDialog>
#include <QPixmap>

class QListWidget;

class InsertChartDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertChartDialog(QWidget *parent = nullptr);

    QString chartType() const;
    QPixmap generateChart(const QString &type) const;

private:
    QListWidget *m_chartList = nullptr;
};
