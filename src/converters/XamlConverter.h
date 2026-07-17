#pragma once

#include <QString>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextBlockFormat>
#include <QTextCharFormat>
#include <QTextBlock>
#include <QTextList>
#include <QTextTable>
#include <QMarginsF>
#include <QColor>

class XamlConverter
{
public:
    static bool loadFromXaml(const QString &xml, QTextDocument *doc,
                             QMarginsF &outMargins, QColor &outPageBackground);
    static QString saveToXaml(const QTextDocument *doc,
                              const QMarginsF &margins,
                              const QColor &pageBackground = QColor());
};
