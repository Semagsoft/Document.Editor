#include "DpiUtils.h"

#include <QApplication>
#include <QScreen>
#include <QWindow>

qreal DpiUtils::screenDpiX()
{
    QWindow *win = QApplication::focusWindow();
    if (win && win->screen())
        return win->screen()->logicalDotsPerInchX();
    return 96.0;
}

qreal DpiUtils::screenDpiY()
{
    QWindow *win = QApplication::focusWindow();
    if (win && win->screen())
        return win->screen()->logicalDotsPerInchY();
    return 96.0;
}

qreal DpiUtils::mmToDip(qreal mm)
{
    return mm * screenDpiX() / 25.4;
}

qreal DpiUtils::cmToDip(qreal cm)
{
    return cm * screenDpiX() / 2.54;
}

qreal DpiUtils::inchToDip(qreal inches)
{
    return inches * screenDpiX();
}

qreal DpiUtils::ptToDip(qreal pt)
{
    return pt * screenDpiX() / 72.0;
}

qreal DpiUtils::dipToInch(qreal dip)
{
    return dip / screenDpiX();
}

qreal DpiUtils::dipToCm(qreal dip)
{
    return dip * 2.54 / screenDpiX();
}

qreal DpiUtils::dipToMm(qreal dip)
{
    return dip * 25.4 / screenDpiX();
}
