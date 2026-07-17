#pragma once

#include <QDialog>

class SymbolPicker;

class InsertSymbolDialog : public QDialog
{
    Q_OBJECT

public:
    explicit InsertSymbolDialog(QWidget *parent = nullptr);

    QChar selectedSymbol() const;

private:
    SymbolPicker *m_symbolPicker = nullptr;
    QChar m_symbol;
};
