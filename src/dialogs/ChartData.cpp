#include "ChartData.h"

ChartData ChartData::defaultData(const QString& type)
{
    ChartData data;
    data.type = type;
    data.title = type;

    ChartSeriesData series;
    series.name = QStringLiteral("Values");
    series.categories = {QStringLiteral("Jan"), QStringLiteral("Feb"), QStringLiteral("Mar"), QStringLiteral("Apr"), QStringLiteral("May")};
    series.values = {5, 8, 12, 6, 10};
    data.series.append(series);

    if (type.contains(QStringLiteral("Pie"), Qt::CaseInsensitive) || type.contains(QStringLiteral("Donut"), Qt::CaseInsensitive)) {
        series.categories = {QStringLiteral("Q1"), QStringLiteral("Q2"), QStringLiteral("Q3"), QStringLiteral("Q4")};
        series.values = {40, 30, 20, 10};
        data.series.clear();
        data.series.append(series);
    }
    return data;
}
