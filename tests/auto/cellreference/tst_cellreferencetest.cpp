#include "xlsxcellreference.h"
#include <QString>
#include <QtTest>

using namespace QXlsx;

class CellReferenceTest : public QObject
{
    Q_OBJECT

public:
    CellReferenceTest();

private Q_SLOTS:
    void test_toString_data();
    void test_toString();
    void test_fromString_data();
    void test_fromString();
};

CellReferenceTest::CellReferenceTest()
{
}

void CellReferenceTest::test_fromString()
{
    QFETCH(QString, cell);
    QFETCH(int, row);
    QFETCH(int, col);
    CellReference pos(cell);
    QCOMPARE(pos.row(), row);
    QCOMPARE(pos.column(), col);
}

void CellReferenceTest::test_fromString_data()
{
    QTest::addColumn<QString>("cell");
    QTest::addColumn<int>("row");
    QTest::addColumn<int>("col");

    QTest::newRow("A1") << "A1" << 1 << 1;
    QTest::newRow("B1") << "B1" << 1 << 2;
    QTest::newRow("C1") << "C1" << 1 << 3;
    QTest::newRow("J1") << "J1" << 1 << 10;
    QTest::newRow("A2") << "A2" << 2 << 1;
    QTest::newRow("A3") << "A3" << 3 << 1;
    QTest::newRow("A10") << "$A10" << 10 << 1;
    QTest::newRow("Z8") << "Z$8" << 8 << 26;
    QTest::newRow("AA10") << "$AA$10" << 10 << 27;
    QTest::newRow("IU2") << "IU2" << 2 << 255;
    QTest::newRow("XFD1") << "XFD1" << 1 << 16384;
    QTest::newRow("XFE1048577") << "XFE1048577" << 1048577 << 16385;
}

void CellReferenceTest::test_toString()
{
    QFETCH(int, row);
    QFETCH(int, col);
    QFETCH(bool, row_abs);
    QFETCH(bool, col_abs);
    QFETCH(QString, cell);

    QCOMPARE(CellReference(row,col).toString(row_abs, col_abs), cell);
}

void CellReferenceTest::test_toString_data()
{
    QTest::addColumn<int>("row");
    QTest::addColumn<int>("col");
    QTest::addColumn<bool>("row_abs");
    QTest::addColumn<bool>("col_abs");
    QTest::addColumn<QString>("cell");

    QTest::newRow("simple") << 1 << 1 << false << false << "A1";
    QTest::newRow("rowabs") << 1 << 1 << true << false << "A$1";
    QTest::newRow("colabs") << 1 << 1 << false << true << "$A1";
    QTest::newRow("bothabs") << 1 << 1 << true << true << "$A$1";
    QTest::newRow("...") << 1048577 << 16385 << false << false << "XFE1048577";
}

QTEST_APPLESS_MAIN(CellReferenceTest)

#include "tst_cellreferencetest.moc"
