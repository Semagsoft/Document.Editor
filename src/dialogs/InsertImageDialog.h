#pragma once

#include <QDialog>

class QLineEdit;
class QLabel;
class QComboBox;

class InsertImageDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertImageDialog(QWidget *parent = nullptr);

    QString imagePath() const;
    int width() const;
    int height() const;

private:
    QLineEdit *m_pathEdit = nullptr;
    QLabel *m_previewLabel = nullptr;
    QComboBox *m_widthCombo = nullptr;
    QComboBox *m_heightCombo = nullptr;

    void browseImage();
    void updatePreview(const QString &path);
};
