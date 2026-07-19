#include "RulerWidget.h"
#include "utils/DpiUtils.h"

#include <QPainter>
#include <QMouseEvent>
#include <QPen>
#include <QFontMetricsF>

RulerWidget::RulerWidget(QWidget *parent)
    : QWidget(parent)
{
    setFixedHeight(static_cast<int>(m_height));
    setMouseTracking(true);
}

RulerWidget::Unit RulerWidget::unit() const { return m_unit; }
void RulerWidget::setUnit(Unit unit) { m_unit = unit; update(); }

qreal RulerWidget::length() const { return m_length; }

void RulerWidget::setLength(qreal length)
{
    m_length = length;
    if (m_autoSize) {
        qreal dipLength = (m_unit == Inches)
            ? DpiUtils::inchToDip(length)
            : DpiUtils::cmToDip(length);
        setFixedWidth(static_cast<int>(dipLength * m_zoom));
    }
    update();
}

qreal RulerWidget::chipPosition() const { return m_chipPos; }

void RulerWidget::setChipPosition(qreal pos)
{
    m_chipPos = pos;
    update();
}

bool RulerWidget::autoSize() const { return m_autoSize; }

void RulerWidget::setAutoSize(bool autoSize)
{
    m_autoSize = autoSize;
    if (autoSize)
        setLength(m_length);
}

qreal RulerWidget::zoom() const { return m_zoom; }

void RulerWidget::setZoom(qreal zoom)
{
    m_zoom = qBound(0.1, zoom, 5.0);
    if (m_autoSize)
        setLength(m_length);
    update();
}

QSize RulerWidget::sizeHint() const
{
    constexpr qreal kDpi = 96.0;
    constexpr qreal kCmPerInch = 2.54;
    return QSize(static_cast<int>(m_length * kDpi / kCmPerInch), static_cast<int>(m_height));
}

void RulerWidget::paintEvent(QPaintEvent *)
{
    QPainter painter(this);
    painter.setRenderHint(QPainter::Antialiasing);

    QRectF rect = this->rect().adjusted(0.5, 0.5, -0.5, -0.5);

    // Background
    painter.fillRect(rect, QColor(248, 248, 248));

    // Border
    painter.setPen(QPen(QColor(180, 180, 180), 1));
    painter.drawRect(rect);

    // Ticks
    painter.save();
    drawTicks(painter, rect);
    painter.restore();

    // Chip marker
    if (m_chipPos > -100)
        drawChip(painter, m_chipPos);
}

void RulerWidget::drawTicks(QPainter &painter, const QRectF &rect)
{
    painter.setPen(QPen(Qt::black, 1));
    qreal dpiFactor = DpiUtils::screenDpiX() / 96.0;

    if (m_unit == Inches) {
        // Major ticks every inch
        qreal inchInPixels = DpiUtils::inchToDip(1.0) * m_zoom;
        qreal totalInches = m_length;
        for (int i = 0; i <= static_cast<int>(totalInches); ++i) {
            qreal x = i * inchInPixels;
            if (x > width()) break;

            // Major tick
            painter.drawLine(QPointF(x, rect.bottom()), QPointF(x, rect.bottom() - m_majorTick));

            // Label
            if (i > 0) {
                QFont labelFont = painter.font();
                labelFont.setPixelSize(9);
                painter.setFont(labelFont);
                painter.drawText(QPointF(x + 2, rect.bottom() - m_majorTick - 1),
                                 QString::number(i));
            }

            // Minor ticks (quarter and half inches)
            if (i < totalInches) {
                // Half inch
                qreal halfX = x + inchInPixels * 0.5;
                if (halfX <= width()) {
                    painter.drawLine(QPointF(halfX, rect.bottom()),
                                     QPointF(halfX, rect.bottom() - m_midTick));
                }

                // Quarter inches
                for (int q = 1; q <= 3; q += 2) {
                    qreal quarterX = x + inchInPixels * q * 0.25;
                    if (quarterX <= width()) {
                        painter.drawLine(QPointF(quarterX, rect.bottom()),
                                         QPointF(quarterX, rect.bottom() - m_minorTick));
                    }
                }
            }
        }
    } else {
        // Centimeter ticks
        qreal cmInPixels = DpiUtils::cmToDip(1.0) * m_zoom;
        qreal totalCm = m_length;
        for (int i = 0; i <= static_cast<int>(totalCm); ++i) {
            qreal x = i * cmInPixels;
            if (x > width()) break;

            painter.drawLine(QPointF(x, rect.bottom()), QPointF(x, rect.bottom() - m_majorTick));

            if (i > 0 && i % 5 == 0) {
                QFont labelFont = painter.font();
                labelFont.setPixelSize(9);
                painter.setFont(labelFont);
                painter.drawText(QPointF(x + 2, rect.bottom() - m_majorTick - 1),
                                 QString::number(i));
            }

            // Millimeter ticks
            if (i < totalCm) {
                for (int mm = 1; mm < 10; ++mm) {
                    qreal mmX = x + cmInPixels * mm / 10.0;
                    if (mmX > width()) break;
                    qreal tickH = (mm == 5) ? m_midTick : m_minorTick;
                    painter.drawLine(QPointF(mmX, rect.bottom()),
                                     QPointF(mmX, rect.bottom() - tickH));
                }
            }
        }
    }
}

void RulerWidget::drawChip(QPainter &painter, qreal x)
{
    painter.setPen(QPen(QColor(220, 40, 40), 1.5));
    painter.drawLine(QPointF(x, 0), QPointF(x, static_cast<qreal>(height())));
}

void RulerWidget::mousePressEvent(QMouseEvent *event)
{
    m_chipPos = event->position().x();
    update();
}

void RulerWidget::mouseMoveEvent(QMouseEvent *event)
{
    if (event->buttons() & Qt::LeftButton) {
        m_chipPos = event->position().x();
        update();
    }
}
