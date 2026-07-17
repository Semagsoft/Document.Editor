#pragma once

#include <QString>
#include <QTextDocument>

class RtfConverter
{
public:
    static bool loadFromRtf(const QByteArray &rtfData, QTextDocument *doc);
    static QByteArray saveToRtf(const QTextDocument *doc);
};
