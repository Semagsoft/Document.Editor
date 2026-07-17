#pragma once

#include <QWidget>
#include <QString>
#include <QList>

class SymbolPicker : public QWidget
{
    Q_OBJECT

public:
    explicit SymbolPicker(QWidget *parent = nullptr);

    QSize sizeHint() const override;

signals:
    void symbolSelected(const QString &symbol);

protected:
    void paintEvent(QPaintEvent *event) override;
    void mouseMoveEvent(QMouseEvent *event) override;
    void leaveEvent(QEvent *event) override;
    void mousePressEvent(QMouseEvent *event) override;

private:
    struct Symbol {
        QString character;
        QString category;
    };

    QList<Symbol> m_symbols;
    int m_columns = 8;
    int m_hoverIndex = -1;

    static constexpr int m_cellSize = 32;
    static constexpr int m_gap = 2;
    static constexpr int m_margin = 4;
    static constexpr int m_headerHeight = 20;
};
