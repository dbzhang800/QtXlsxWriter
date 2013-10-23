/****************************************************************************
** Copyright (c) 2013 Debao Zhang <hello@debao.me>
** All right reserved.
**
** Permission is hereby granted, free of charge, to any person obtaining
** a copy of this software and associated documentation files (the
** "Software"), to deal in the Software without restriction, including
** without limitation the rights to use, copy, modify, merge, publish,
** distribute, sublicense, and/or sell copies of the Software, and to
** permit persons to whom the Software is furnished to do so, subject to
** the following conditions:
**
** The above copyright notice and this permission notice shall be
** included in all copies or substantial portions of the Software.
**
** THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
** EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
** MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
** NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
** LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
** OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
** WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
**
****************************************************************************/
#include "private/xlsxutility_p.h"
#include <QString>
#include <QtTest>
#include <QDateTime>

class UtilityTest : public QObject
{
    Q_OBJECT
    
public:
    UtilityTest();
    
private Q_SLOTS:
    void test_cell_to_rowcol();
    void test_cell_to_rowcol_data();

    void test_rowcol_to_cell();
    void test_rowcol_to_cell_data();

    void test_datetimeToNumber_data();
    void test_datetimeToNumber();

    void test_datetimeFromNumber_data();
    void test_datetimeFromNumber();
};

UtilityTest::UtilityTest()
{
}

void UtilityTest::test_cell_to_rowcol()
{
    QFETCH(QString, cell);
    QFETCH(int, row);
    QFETCH(int, col);
    QPoint pos = QXlsx::xl_cell_to_rowcol(cell);
    QCOMPARE(pos.x(), row);
    QCOMPARE(pos.y(), col);
}

void UtilityTest::test_cell_to_rowcol_data()
{
    QTest::addColumn<QString>("cell");
    QTest::addColumn<int>("row");
    QTest::addColumn<int>("col");

    QTest::newRow("A1") << "A1" << 0 << 0;
    QTest::newRow("B1") << "B1" << 0 << 1;
    QTest::newRow("C1") << "C1" << 0 << 2;
    QTest::newRow("J1") << "J1" << 0 << 9;
    QTest::newRow("A2") << "A2" << 1 << 0;
    QTest::newRow("A3") << "A3" << 2 << 0;
    QTest::newRow("A10") << "A10" << 9 << 0;
    QTest::newRow("Z8") << "Z8" << 7 << 25;
    QTest::newRow("AA10") << "AA10" << 9 << 26;
    QTest::newRow("IU2") << "IU2" << 1 << 254;
    QTest::newRow("XFD1") << "XFD1" << 0 << 16383;
    QTest::newRow("XFE1048577") << "XFE1048577" << 1048576 << 16384;
}

void UtilityTest::test_rowcol_to_cell()
{
    QFETCH(int, row);
    QFETCH(int, col);
    QFETCH(bool, row_abs);
    QFETCH(bool, col_abs);
    QFETCH(QString, cell);

    QCOMPARE(QXlsx::xl_rowcol_to_cell(row, col, row_abs, col_abs), cell);
}

void UtilityTest::test_rowcol_to_cell_data()
{
    QTest::addColumn<int>("row");
    QTest::addColumn<int>("col");
    QTest::addColumn<bool>("row_abs");
    QTest::addColumn<bool>("col_abs");
    QTest::addColumn<QString>("cell");

    QTest::newRow("simple") << 0 << 0 << false << false << "A1";
    QTest::newRow("rowabs") << 0 << 0 << true << false << "A$1";
    QTest::newRow("colabs") << 0 << 0 << false << true << "$A1";
    QTest::newRow("bothabs") << 0 << 0 << true << true << "$A$1";
    QTest::newRow("...") << 1048576 << 16384 << false << false << "XFE1048577";
}

void UtilityTest::test_datetimeToNumber_data()
{
    QTest::addColumn<QDateTime>("dt");
    QTest::addColumn<bool>("is1904");
    QTest::addColumn<double>("num");

    //Note, for number 0, Excel2007 shown as 1900-1-0, which should be 1899-12-31
    QTest::newRow("0") << QDateTime(QDate(1899, 12, 31), QTime(0,0), Qt::UTC) << false << 0.0;
    QTest::newRow("1.25") << QDateTime(QDate(1900, 1, 1), QTime(6, 0), Qt::UTC) << false << 1.25;
    QTest::newRow("59") << QDateTime(QDate(1900, 2, 28), QTime(0, 0), Qt::UTC) << false << 59.0;
    QTest::newRow("61") << QDateTime(QDate(1900, 3, 1), QTime(0, 0), Qt::UTC) << false << 61.0;

    QTest::newRow("1904: 0") << QDateTime(QDate(1904, 1, 1), QTime(0,0), Qt::UTC) << true << 0.0;
    QTest::newRow("1904: 1.25") << QDateTime(QDate(1904, 1, 2), QTime(6, 0), Qt::UTC) << true << 1.25;
}

void UtilityTest::test_datetimeToNumber()
{
    QFETCH(QDateTime, dt);
    QFETCH(bool, is1904);
    QFETCH(double, num);

    QCOMPARE(QXlsx::datetimeToNumber(dt, is1904), num);
}

void UtilityTest::test_datetimeFromNumber_data()
{
    QTest::addColumn<QDateTime>("dt");
    QTest::addColumn<bool>("is1904");
    QTest::addColumn<double>("num");

    QTest::newRow("0") << QDateTime(QDate(1899, 12, 31), QTime(0,0), Qt::UTC) << false << 0.0;
    QTest::newRow("1.25") << QDateTime(QDate(1900, 1, 1), QTime(6, 0), Qt::UTC) << false << 1.25;
    QTest::newRow("59") << QDateTime(QDate(1900, 2, 28), QTime(0,0), Qt::UTC) << false << 59.0;
    QTest::newRow("61") << QDateTime(QDate(1900, 3, 1), QTime(0,0), Qt::UTC) << false << 61.0;

    QTest::newRow("1904: 0") << QDateTime(QDate(1904, 1, 1), QTime(0,0), Qt::UTC) << true << 0.0;
    QTest::newRow("1904: 1.25") << QDateTime(QDate(1904, 1, 2), QTime(6, 0), Qt::UTC) << true << 1.25;
}

void UtilityTest::test_datetimeFromNumber()
{
    QFETCH(QDateTime, dt);
    QFETCH(bool, is1904);
    QFETCH(double, num);

    QCOMPARE(QXlsx::datetimeFromNumber(num, is1904), dt);
}

QTEST_APPLESS_MAIN(UtilityTest)

#include "tst_utilitytest.moc"
