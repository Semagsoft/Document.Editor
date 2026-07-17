#include "TableGridPicker.h"

#include <QPainter>
#include <QMouseEvent>
#include <QToolTip>

TableGridPicker::TableGridPicker(QWidget *parent)
    : QWidget(parent)
{
    setMouseTracking(true);
    int w = m_margin * 2 + m_maxCols * (m_cellSize + m_gap) - m_gap;
    int h = m_margin * 2 + m_maxRows * (m_cellSize + m_gap) - m_gap;
    setFixedSize(w, h + 20);
}

QSize TableGridPicker::sizeHint() const
{
    return QSize(m_margin * 2 + m_maxCols * (m_cellSize + m_gap) - m_gap,
                 m_margin * 2 + m_maxRows * (m_cellSize + m_gap) - m_gap + 20);
}

void TableGridPicker::paintEvent(QPaintEvent *)
{
    QPainter painter(this);
    painter.setRenderHint(QPainter::Antialiasing);

    // Background
    painter.fillRect(rect(), QColor(255, 255, 255));

    // Grid cells
    for (int row = 0; row < m_maxRows; ++row) {
        for (int col = 0; col < m_maxCols; ++col) {
            int x = m_margin + col * (m_cellSize + m_gap);
            int y = m_margin + row * (m_cellSize + m_gap);

            QRect cellRect(x, y, m_cellSize, m_cellSize);

            bool highlighted = (row <= m_highlightRow && col <= m_highlightCol);

            if (highlighted) {
                painter.setBrush(QColor(180, 205, 240));
                painter.setPen(QPen(QColor(60, 120, 200), 1.5));
            } else {
                painter.setBrush(QColor(240, 240, 240));
                painter.setPen(QPen(QColor(200, 200, 200), 1));
            }

            painter.drawRect(cellRect);
        }
    }

    // Label
    if (m_highlightRow >= 0 && m_highlightCol >= 0) {
        painter.setPen(Qt::black);
        QFont f = painter.font();
        f.setPixelSize(11);
        painter.setFont(f);
        QString label = QStringLiteral("%1 x %2 Table")
                            .arg(m_highlightRow + 1).arg(m_highlightCol + 1);
        painter.drawText(QRect(0, height() - 18, width(), 18),
                         Qt::AlignCenter, label);
    }
}

void TableGridPicker::mouseMoveEvent(QMouseEvent *event)
{
    QPoint pos = event->position().toPoint();
    int col = (pos.x() - m_margin) / (m_cellSize + m_gap);
    int row = (pos.y() - m_margin) / (m_cellSize + m_gap);

    if (row < 0) row = 0;
    if (col < 0) col = 0;
    if (row >= m_maxRows) row = m_maxRows - 1;
    if (col >= m_maxCols) col = m_maxCols - 1;

    if (row != m_highlightRow || col != m_highlightCol) {
        m_highlightRow = row;
        m_highlightCol = col;
        update();
    }
}

void TableGridPicker::leaveEvent(QEvent *)
{
    m_highlightRow = -1;
    m_highlightCol = -1;
    update();
}

void TableGridPicker::mousePressEvent(QMouseEvent *event)
{
    if (event->button() == Qt::LeftButton && m_highlightRow >= 0 && m_highlightCol >= 0) {
        emit tableSelected(m_highlightRow + 1, m_highlightCol + 1);
    }
}
