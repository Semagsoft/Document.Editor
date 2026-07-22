#pragma once

#include <QString>
#include <QByteArray>
#include <QTextDocument>
#include <QMarginsF>
#include <QColor>

class DocxConverter
{
public:
    static bool loadFromDocx(const QByteArray &zipData, QTextDocument *doc,
                             QMarginsF &outMargins, QColor &outPageBackground);
    static QByteArray saveToDocx(const QTextDocument *doc,
                                 const QMarginsF &margins,
                                 const QColor &pageBackground = QColor());
};
