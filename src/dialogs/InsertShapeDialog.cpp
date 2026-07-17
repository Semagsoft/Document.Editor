#include "InsertShapeDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QListWidget>
#include <QPushButton>
#include <QPainter>
#include <QPainterPath>
#include <QPolygonF>

InsertShapeDialog::InsertShapeDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Insert Shape"));
    setFixedSize(280, 280);

    auto *layout = new QVBoxLayout(this);

    layout->addWidget(new QLabel(tr("Select shape:")));

    m_shapeList = new QListWidget();
    QStringList shapes = {
        tr("Rectangle"),
        tr("Rounded Rectangle"),
        tr("Ellipse"),
        tr("Circle"),
        tr("Triangle"),
        tr("Diamond"),
        tr("Arrow Right"),
        tr("Arrow Left"),
        tr("Star"),
        tr("Hexagon"),
    };
    m_shapeList->addItems(shapes);
    m_shapeList->setCurrentRow(0);
    layout->addWidget(m_shapeList);

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

QString InsertShapeDialog::shapeName() const
{
    return m_shapeList->currentItem() ? m_shapeList->currentItem()->text() : QString();
}

QPixmap InsertShapeDialog::generateShape(const QString &name) const
{
    const int size = 200;
    QPixmap pixmap(size, size);
    pixmap.fill(Qt::transparent);
    QPainter painter(&pixmap);
    painter.setRenderHint(QPainter::Antialiasing);
    painter.setPen(QPen(Qt::black, 3));
    painter.setBrush(QColor(66, 133, 244));
    QRectF rect(20, 20, size - 40, size - 40);

    if (name == tr("Rectangle")) {
        painter.drawRect(rect);
    } else if (name == tr("Rounded Rectangle")) {
        painter.drawRoundedRect(rect, 20, 20);
    } else if (name == tr("Ellipse")) {
        painter.drawEllipse(rect);
    } else if (name == tr("Circle")) {
        QRectF square(30, 30, size - 60, size - 60);
        painter.drawEllipse(square);
    } else if (name == tr("Triangle")) {
        QPolygonF triangle;
        triangle << QPointF(size / 2.0, 20) << QPointF(20, size - 20) << QPointF(size - 20, size - 20);
        painter.drawPolygon(triangle);
    } else if (name == tr("Diamond")) {
        QPolygonF diamond;
        diamond << QPointF(size / 2.0, 20) << QPointF(size - 20, size / 2.0)
                << QPointF(size / 2.0, size - 20) << QPointF(20, size / 2.0);
        painter.drawPolygon(diamond);
    } else if (name == tr("Arrow Right")) {
        QPolygonF arrow;
        arrow << QPointF(20, 30) << QPointF(size * 0.6, 30) << QPointF(size * 0.6, 20)
              << QPointF(size - 20, size / 2.0) << QPointF(size * 0.6, size - 20)
              << QPointF(size * 0.6, size - 30) << QPointF(20, size - 30);
        painter.drawPolygon(arrow);
    } else if (name == tr("Arrow Left")) {
        QPolygonF arrow;
        arrow << QPointF(size - 20, 30) << QPointF(size * 0.4, 30) << QPointF(size * 0.4, 20)
              << QPointF(20, size / 2.0) << QPointF(size * 0.4, size - 20)
              << QPointF(size * 0.4, size - 30) << QPointF(size - 20, size - 30);
        painter.drawPolygon(arrow);
    } else if (name == tr("Star")) {
        QPainterPath starPath;
        starPath.moveTo(size / 2.0, 10);
        for (int i = 1; i < 10; ++i) {
            qreal angle = M_PI / 2.0 - i * M_PI / 5.0;
            qreal r = (i % 2 == 0) ? 90.0 : 40.0;
            starPath.lineTo(size / 2.0 + r * qCos(angle), size / 2.0 - r * qSin(angle));
        }
        starPath.closeSubpath();
        painter.drawPath(starPath);
    } else if (name == tr("Hexagon")) {
        QPolygonF hex;
        for (int i = 0; i < 6; ++i) {
            qreal angle = M_PI / 3.0 * i - M_PI / 6.0;
            hex << QPointF(size / 2.0 + (size / 2.0 - 20) * qCos(angle),
                           size / 2.0 + (size / 2.0 - 20) * qSin(angle));
        }
        painter.drawPolygon(hex);
    }

    painter.end();
    return pixmap;
}
