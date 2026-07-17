#pragma once

#include <QtGlobal>

namespace DpiUtils {

qreal screenDpiX();
qreal screenDpiY();

qreal mmToDip(qreal mm);
qreal cmToDip(qreal cm);
qreal inchToDip(qreal inches);
qreal ptToDip(qreal pt);

qreal dipToInch(qreal dip);
qreal dipToCm(qreal dip);
qreal dipToMm(qreal dip);

} // namespace DpiUtils
