#pragma once

#include <QWidget>
#include <QSize>

class QGridLayout;

class TableGridPicker : public QWidget
{
    Q_OBJECT

public:
    explicit TableGridPicker(QWidget *parent = nullptr);

    QSize sizeHint() const override;

signals:
    void tableSelected(int rows, int cols);

protected:
    void paintEvent(QPaintEvent *event) override;
    void mouseMoveEvent(QMouseEvent *event) override;
    void leaveEvent(QEvent *event) override;
    void mousePressEvent(QMouseEvent *event) override;

private:
    void updateHighlight(int row, int col);

    static constexpr int m_maxRows = 8;
    static constexpr int m_maxCols = 10;
    static constexpr int m_cellSize = 18;
    static constexpr int m_gap = 2;
    static constexpr int m_margin = 8;

    int m_highlightRow = -1;
    int m_highlightCol = -1;
};
