#pragma once

#include <QDialog>
#include <QPixmap>

class QListWidget;

class InsertShapeDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertShapeDialog(QWidget *parent = nullptr);

    QString shapeName() const;
    QPixmap generateShape(const QString &name) const;

private:
    QListWidget *m_shapeList = nullptr;
};
