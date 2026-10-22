#pragma once

#include <QDialog>
#include <QPixmap>

#include "ChartData.h"

class QListWidget;

class InsertChartDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertChartDialog(QWidget *parent = nullptr);

    QString chartType() const;
    ChartData chartData() const;
    QPixmap generateChart(const QString &type) const;

private:
    QListWidget *m_chartList = nullptr;
};
