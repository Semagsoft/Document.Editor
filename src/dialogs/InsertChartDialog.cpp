#include "InsertChartDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QListWidget>
#include <QPushButton>
#include <QPainter>

#include <QtCharts/QChart>
#include <QtCharts/QChartView>
#include <QtCharts/QPieSeries>
#include <QtCharts/QBarSeries>
#include <QtCharts/QBarSet>
#include <QtCharts/QBarCategoryAxis>
#include <QtCharts/QValueAxis>
#include <QtCharts/QLineSeries>
#include <QtCharts/QCategoryAxis>

InsertChartDialog::InsertChartDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Chart"));
    setFixedSize(280, 300);

    auto *layout = new QVBoxLayout(this);

    layout->addWidget(new QLabel(tr("Select chart type:")));

    m_chartList = new QListWidget();
    QStringList charts = {
        tr("Bar Chart"),
        tr("Line Chart"),
        tr("Pie Chart"),
        tr("Spline Chart"),
        tr("Scatter Chart"),
        tr("Area Chart"),
        tr("Column Chart"),
        tr("Donut Chart"),
    };
    m_chartList->addItems(charts);
    m_chartList->setCurrentRow(0);
    layout->addWidget(m_chartList);

    auto *btnLayout = new QHBoxLayout();
    auto *okBtn = new QPushButton(tr("Insert"));
    auto *cancelBtn = new QPushButton(tr("Cancel"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

QString InsertChartDialog::chartType() const
{
    return m_chartList->currentItem() ? m_chartList->currentItem()->text() : QString();
}

static int chartTypeIndex(const QString &type)
{
    if (type == QObject::tr("Pie Chart") || type == QObject::tr("Donut Chart"))
        return 0;
    if (type == QObject::tr("Bar Chart") || type == QObject::tr("Column Chart"))
        return 1;
    return 2;
}

QPixmap InsertChartDialog::generateChart(const QString &type) const
{
    const int w = 400, h = 300;

    auto *chart = new QChart();
    chart->setTitle(type);
    chart->setAnimationOptions(QChart::SeriesAnimations);
    chart->legend()->setAlignment(Qt::AlignBottom);

    switch (chartTypeIndex(type)) {
    case 0: {
        auto *series = new QPieSeries();
        series->append(QStringLiteral("Q1"), 40);
        series->append(QStringLiteral("Q2"), 30);
        series->append(QStringLiteral("Q3"), 20);
        series->append(QStringLiteral("Q4"), 10);
        chart->addSeries(series);
        break;
    }
    case 1: {
        auto *set = new QBarSet(QStringLiteral("Values"));
        *set << 5 << 8 << 12 << 6 << 10;
        auto *series = new QBarSeries();
        series->append(set);
        chart->addSeries(series);
        auto *axisX = new QBarCategoryAxis();
        axisX->append(QStringList{ QStringLiteral("Jan"), QStringLiteral("Feb"),
            QStringLiteral("Mar"), QStringLiteral("Apr"), QStringLiteral("May") });
        chart->addAxis(axisX, Qt::AlignBottom);
        series->attachAxis(axisX);
        auto *axisY = new QValueAxis();
        axisY->setRange(0, 15);
        chart->addAxis(axisY, Qt::AlignLeft);
        series->attachAxis(axisY);
        break;
    }
    default: {
        auto *series = new QLineSeries();
        series->append(0, 5);
        series->append(1, 8);
        series->append(2, 12);
        series->append(3, 6);
        series->append(4, 10);
        chart->addSeries(series);
        auto *axisX = new QCategoryAxis();
        axisX->append(QStringLiteral("Jan"), 0);
        axisX->append(QStringLiteral("Feb"), 1);
        axisX->append(QStringLiteral("Mar"), 2);
        axisX->append(QStringLiteral("Apr"), 3);
        axisX->append(QStringLiteral("May"), 4);
        axisX->setRange(-0.5, 4.5);
        chart->addAxis(axisX, Qt::AlignBottom);
        series->attachAxis(axisX);
        auto *axisY = new QValueAxis();
        axisY->setRange(0, 15);
        chart->addAxis(axisY, Qt::AlignLeft);
        series->attachAxis(axisY);
        break;
    }
    }

    QChartView view(chart);
    view.resize(w, h);

    QPixmap pixmap(w, h);
    pixmap.fill(Qt::white);
    QPainter painter(&pixmap);
    painter.setRenderHint(QPainter::Antialiasing);
    view.render(&painter);
    painter.end();

    return pixmap;
}
