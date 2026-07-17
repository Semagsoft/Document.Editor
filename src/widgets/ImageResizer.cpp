#include "ImageResizer.h"

#include <QPainter>
#include <QMouseEvent>
#include <QPen>

ImageResizer::ImageResizer(QWidget *parent)
    : QWidget(parent)
{
    setFixedSize(100, 100);
    updateHandles();
}

void ImageResizer::setImageSize(const QSize &size)
{
    m_imageSize = size;
    setFixedSize(size);
    updateHandles();
    update();
}

QSize ImageResizer::imageSize() const
{
    return m_imageSize;
}

void ImageResizer::paintEvent(QPaintEvent *)
{
    QPainter painter(this);

    // Draw border around image
    painter.setPen(QPen(QColor(0, 120, 215), 2));
    painter.setBrush(Qt::NoBrush);
    painter.drawRect(rect().adjusted(1, 1, -1, -1));

    // Draw handles
    painter.setBrush(QColor(255, 255, 255));
    painter.setPen(QPen(QColor(0, 120, 215), 1));

    for (int i = 0; i < 4; ++i) {
        painter.drawRect(m_handles[i]);
    }
}

void ImageResizer::mousePressEvent(QMouseEvent *event)
{
    if (event->button() == Qt::LeftButton) {
        m_activeHandle = handleAtPos(event->position().toPoint());
        if (m_activeHandle != None) {
            m_dragging = true;
            m_dragStart = event->position().toPoint();
            m_startSize = m_imageSize;
        }
    }
}

void ImageResizer::mouseMoveEvent(QMouseEvent *event)
{
    if (m_dragging && m_activeHandle != None) {
        QPoint delta = event->position().toPoint() - m_dragStart;
        QSize newSize = m_startSize;

        switch (m_activeHandle) {
        case BottomRight:
            newSize = QSize(qMax(20, m_startSize.width() + delta.x()),
                            qMax(20, m_startSize.height() + delta.y()));
            break;
        case BottomLeft:
            newSize = QSize(qMax(20, m_startSize.width() - delta.x()),
                            qMax(20, m_startSize.height() + delta.y()));
            break;
        case TopRight:
            newSize = QSize(qMax(20, m_startSize.width() + delta.x()),
                            qMax(20, m_startSize.height() - delta.y()));
            break;
        case TopLeft:
            newSize = QSize(qMax(20, m_startSize.width() - delta.x()),
                            qMax(20, m_startSize.height() - delta.y()));
            break;
        default:
            break;
        }

        setImageSize(newSize);
        emit sizeChanged(newSize);
    }
}

void ImageResizer::mouseReleaseEvent(QMouseEvent *)
{
    m_dragging = false;
    m_activeHandle = None;
}

ImageResizer::Handle ImageResizer::handleAtPos(const QPoint &pos) const
{
    for (int i = 0; i < 4; ++i) {
        if (m_handles[i].contains(pos))
            return static_cast<Handle>(i);
    }
    return None;
}

QRect ImageResizer::handleRect(Handle handle) const
{
    int hs = m_handleSize;
    int w = m_imageSize.width();
    int h = m_imageSize.height();

    switch (handle) {
    case TopLeft:     return QRect(0, 0, hs, hs);
    case TopRight:    return QRect(w - hs, 0, hs, hs);
    case BottomLeft:  return QRect(0, h - hs, hs, hs);
    case BottomRight: return QRect(w - hs, h - hs, hs, hs);
    default:          return QRect();
    }
}

void ImageResizer::updateHandles()
{
    m_handles[TopLeft]     = handleRect(TopLeft);
    m_handles[TopRight]    = handleRect(TopRight);
    m_handles[BottomLeft]  = handleRect(BottomLeft);
    m_handles[BottomRight] = handleRect(BottomRight);
}
