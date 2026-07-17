#pragma once

#include <QWidget>
#include <QPen>
#include <QFont>

class RulerWidget : public QWidget
{
    Q_OBJECT

public:
    enum Unit { Inches, Centimeters };

    explicit RulerWidget(QWidget *parent = nullptr);

    Unit unit() const;
    void setUnit(Unit unit);

    qreal length() const;
    void setLength(qreal length);

    qreal chipPosition() const;
    void setChipPosition(qreal pos);

    bool autoSize() const;
    void setAutoSize(bool autoSize);

    qreal zoom() const;
    void setZoom(qreal zoom);

    QSize sizeHint() const override;

protected:
    void paintEvent(QPaintEvent *event) override;
    void mousePressEvent(QMouseEvent *event) override;
    void mouseMoveEvent(QMouseEvent *event) override;

private:
    void drawTicks(QPainter &painter, const QRectF &rect);
    void drawChip(QPainter &painter, qreal x);

    Unit m_unit = Inches;
    qreal m_length = 8.5;
    qreal m_chipPos = -1000;
    bool m_autoSize = true;
    qreal m_zoom = 1.0;

    static constexpr qreal m_height = 24.0;
    static constexpr qreal m_minorTick = 4.0;
    static constexpr qreal m_midTick = 8.0;
    static constexpr qreal m_majorTick = 14.0;
};
