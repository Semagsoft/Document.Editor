#pragma once

#include <QDialog>

class QDoubleSpinBox;
class QListWidget;

class LineSpacingDialog : public QDialog
{
    Q_OBJECT

public:
    explicit LineSpacingDialog(qreal currentSpacing, QWidget *parent = nullptr);

    qreal lineSpacing() const;

private:
    QDoubleSpinBox *m_spacingSpin = nullptr;
    QListWidget *m_presetList = nullptr;
    qreal m_customSpacing = 1.0;

    void applyPreset(qreal spacing);
};
