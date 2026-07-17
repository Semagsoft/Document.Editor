#include <QTest>
#include <QApplication>
#include "utils/DpiUtils.h"

class TestDpiUtils : public QObject
{
    Q_OBJECT

private slots:
    void testInchToDip()
    {
        // At 96 DPI, 1 inch = 96 dip
        qreal dpi = DpiUtils::screenDpiX();
        QCOMPARE(DpiUtils::inchToDip(1.0), dpi);
        QCOMPARE(DpiUtils::inchToDip(0.0), 0.0);
        QCOMPARE(DpiUtils::inchToDip(2.0), dpi * 2);
    }

    void testDipToInch()
    {
        qreal dpi = DpiUtils::screenDpiX();
        QCOMPARE(DpiUtils::dipToInch(dpi), 1.0);
        QCOMPARE(DpiUtils::dipToInch(0.0), 0.0);
    }

    void testCmToDip()
    {
        qreal dpi = DpiUtils::screenDpiX();
        // 1 cm = dpi / 2.54
        QCOMPARE(DpiUtils::cmToDip(1.0), dpi / 2.54);
        QCOMPARE(DpiUtils::cmToDip(0.0), 0.0);
    }

    void testDipToCm()
    {
        qreal dpi = DpiUtils::screenDpiX();
        QCOMPARE(DpiUtils::dipToCm(dpi), 2.54);
        QCOMPARE(DpiUtils::dipToCm(0.0), 0.0);
    }

    void testMmToDip()
    {
        qreal dpi = DpiUtils::screenDpiX();
        QCOMPARE(DpiUtils::mmToDip(1.0), dpi / 25.4);
        QCOMPARE(DpiUtils::mmToDip(0.0), 0.0);
    }

    void testDipToMm()
    {
        qreal dpi = DpiUtils::screenDpiX();
        QCOMPARE(DpiUtils::dipToMm(dpi), 25.4);
        QCOMPARE(DpiUtils::dipToMm(0.0), 0.0);
    }

    void testPtToDip()
    {
        qreal dpi = DpiUtils::screenDpiX();
        QCOMPARE(DpiUtils::ptToDip(72.0), dpi);
        QCOMPARE(DpiUtils::ptToDip(0.0), 0.0);
    }

    void testRoundTrip()
    {
        qreal original = 100.0;
        qreal inches = DpiUtils::dipToInch(original);
        qreal back = DpiUtils::inchToDip(inches);
        QCOMPARE(back, original);
    }
};

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    TestDpiUtils test;
    return QTest::qExec(&test, argc, argv);
}
#include "test_DpiUtils.moc"
