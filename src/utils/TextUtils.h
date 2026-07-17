#pragma once

#include <QString>
#include <QStringList>

namespace TextUtils {

QString toUpperCase(const QString &text);
QString toLowerCase(const QString &text);
QString titleCase(const QString &text);
QString sentenceCase(const QString &text);
QString toggleCase(const QString &text);

int wordCount(const QString &text);
int characterCount(const QString &text, bool includeSpaces = true);

} // namespace TextUtils
