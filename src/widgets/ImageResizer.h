#pragma once

#include <QWidget>
#include <QPoint>
#include <QSize>

class ImageResizer : public QWidget
{
    Q_OBJECT

public:
    explicit ImageResizer(QWidget *parent = nullptr);

    void setImageSize(const QSize &size);
    QSize imageSize() const;

signals:
    void sizeChanged(const QSize &newSize);

protected:
    void paintEvent(QPaintEvent *event) override;
    void mousePressEvent(QMouseEvent *event) override;
    void mouseMoveEvent(QMouseEvent *event) override;
    void mouseReleaseEvent(QMouseEvent *event) override;

private:
    enum Handle { None, TopLeft, TopRight, BottomLeft, BottomRight };

    QRect handleRect(Handle handle) const;
    Handle handleAtPos(const QPoint &pos) const;
    void updateHandles();

    static constexpr int m_handleSize = 8;

    QSize m_imageSize;
    Handle m_activeHandle = None;
    QPoint m_dragStart;
    QSize m_startSize;
    QRect m_handles[4];
    bool m_dragging = false;
};
