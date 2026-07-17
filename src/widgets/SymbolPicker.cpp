#include "SymbolPicker.h"

#include <QPainter>
#include <QMouseEvent>

SymbolPicker::SymbolPicker(QWidget *parent)
    : QWidget(parent)
{
    setMouseTracking(true);

    // Currency
    m_symbols.append({QStringLiteral("$"),   tr("Currency")});
    m_symbols.append({QStringLiteral("\u00A2"), tr("Currency")}); // Cent
    m_symbols.append({QStringLiteral("\u00A3"), tr("Currency")}); // Pound
    m_symbols.append({QStringLiteral("\u00A5"), tr("Currency")}); // Yen
    m_symbols.append({QStringLiteral("\u20AC"), tr("Currency")}); // Euro
    m_symbols.append({QStringLiteral("\u00A4"), tr("Currency")}); // Currency sign

    // Misc
    m_symbols.append({QStringLiteral("\u00A9"), tr("Misc")});    // Copyright
    m_symbols.append({QStringLiteral("\u00AE"), tr("Misc")});    // Registered
    m_symbols.append({QStringLiteral("\u2122"), tr("Misc")});    // TM
    m_symbols.append({QStringLiteral("\u00B0"), tr("Misc")});    // Degree
    m_symbols.append({QStringLiteral("\u00B1"), tr("Misc")});    // Plus-minus
    m_symbols.append({QStringLiteral("\u00B7"), tr("Misc")});    // Middle dot

    // Superscripts
    m_symbols.append({QStringLiteral("\u00B9"), tr("Superscript")}); // ¹
    m_symbols.append({QStringLiteral("\u00B2"), tr("Superscript")}); // ²
    m_symbols.append({QStringLiteral("\u00B3"), tr("Superscript")}); // ³

    // Math
    m_symbols.append({QStringLiteral("\u00D7"), tr("Math")});    // ×
    m_symbols.append({QStringLiteral("\u00F7"), tr("Math")});    // ÷
    m_symbols.append({QStringLiteral("\u00BC"), tr("Math")});    // ¼
    m_symbols.append({QStringLiteral("\u00BD"), tr("Math")});    // ½
    m_symbols.append({QStringLiteral("\u00BE"), tr("Math")});    // ¾

    // Arrows
    m_symbols.append({QStringLiteral("\u2190"), tr("Arrows")});  // ←
    m_symbols.append({QStringLiteral("\u2191"), tr("Arrows")});  // ↑
    m_symbols.append({QStringLiteral("\u2192"), tr("Arrows")});  // →
    m_symbols.append({QStringLiteral("\u2193"), tr("Arrows")});  // ↓

    int totalSymbols = m_symbols.size();
    int rows = (totalSymbols / m_columns) + 1 + 1; // +1 for header
    int totalHeight = m_margin * 2 + rows * (m_cellSize + m_gap);
    int totalWidth = m_margin * 2 + m_columns * (m_cellSize + m_gap);
    setFixedSize(totalWidth, totalHeight);
}

QSize SymbolPicker::sizeHint() const
{
    int rows = (m_symbols.size() / m_columns) + 2;
    return QSize(m_margin * 2 + m_columns * (m_cellSize + m_gap),
                 m_margin * 2 + rows * (m_cellSize + m_gap));
}

void SymbolPicker::paintEvent(QPaintEvent *)
{
    QPainter painter(this);
    painter.setRenderHint(QPainter::Antialiasing);

    painter.fillRect(rect(), QColor(255, 255, 255));

    QFont symbolFont = painter.font();
    symbolFont.setPixelSize(16);
    painter.setFont(symbolFont);

    for (int i = 0; i < m_symbols.size(); ++i) {
        int col = i % m_columns;
        int row = i / m_columns;
        int x = m_margin + col * (m_cellSize + m_gap);
        int y = m_margin + row * (m_cellSize + m_gap);

        QRect cellRect(x, y, m_cellSize, m_cellSize);

        if (i == m_hoverIndex) {
            painter.setBrush(QColor(180, 205, 240));
            painter.setPen(QPen(QColor(60, 120, 200), 1));
            painter.drawRect(cellRect);
        }

        painter.setPen(Qt::black);
        painter.drawText(cellRect, Qt::AlignCenter, m_symbols[i].character);
    }
}

void SymbolPicker::mouseMoveEvent(QMouseEvent *event)
{
    QPoint pos = event->position().toPoint();
    int col = (pos.x() - m_margin) / (m_cellSize + m_gap);
    int row = (pos.y() - m_margin) / (m_cellSize + m_gap);

    int index = row * m_columns + col;
    if (index < 0 || index >= m_symbols.size())
        index = -1;

    if (index != m_hoverIndex) {
        m_hoverIndex = index;
        update();
    }
}

void SymbolPicker::leaveEvent(QEvent *)
{
    m_hoverIndex = -1;
    update();
}

void SymbolPicker::mousePressEvent(QMouseEvent *event)
{
    if (event->button() == Qt::LeftButton && m_hoverIndex >= 0) {
        emit symbolSelected(m_symbols[m_hoverIndex].character);
    }
}
