#pragma once

#include <QString>
#include <QStringList>
#include <QVariantList>

struct ChartSeriesData {
    QString name;
    QStringList categories; // for bar/line/pie labels
    QVariantList values;    // numeric values
};

struct ChartData {
    QString type;  // "Bar", "Line", "Pie", "Spline", "Scatter", "Area", "Column", "Donut"
    QString title;
    QList<ChartSeriesData> series;
    // Default demo data for quick insertion
    static ChartData defaultData(const QString& type);
};
