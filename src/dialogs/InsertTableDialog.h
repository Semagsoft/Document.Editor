#pragma once

#include <QDialog>

class TableGridPicker;

class InsertTableDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertTableDialog(QWidget *parent = nullptr);

    int rows() const;
    int columns() const;

private:
    TableGridPicker *m_gridPicker = nullptr;
    int m_rows = 2;
    int m_cols = 2;
};
