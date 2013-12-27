#include "xlsxformat.h"
#include <QString>
#include <QtTest>

QTXLSX_USE_NAMESPACE

class FormatTest : public QObject
{
    Q_OBJECT

public:
    FormatTest();

private Q_SLOTS:
    void testDateTimeFormat();
    void testDateTimeFormat_data();
};

FormatTest::FormatTest()
{
}

void FormatTest::testDateTimeFormat()
{
    QFETCH(QString, data);
    QFETCH(bool, res);

    Format fmt;
    fmt.setNumberFormat(data);

    QCOMPARE(fmt.isDateTimeFormat(), res);
}

void FormatTest::testDateTimeFormat_data()
{
    QTest::addColumn<QString>("data");
    QTest::addColumn<bool>("res");

    QTest::newRow("0") << QString("yyyy-mm-dd")<<true;
    QTest::newRow("1") << QString("m/d/yy")<<true;
    QTest::newRow("2") << QString("h:mm AM/PM")<<true;
    QTest::newRow("3") << QString("m/d/yy h:mm")<<true;
    QTest::newRow("4") << QString("[h]:mm:ss")<<true;
    QTest::newRow("5") << QString("[h]")<<true;
    QTest::newRow("6") << QString("[m]")<<true;
    QTest::newRow("7") << QString("yyyy-mm-dd;###;\\(0.000\\)")<<true;
    QTest::newRow("8") << QString("[Red][m]")<<true;

    QTest::newRow("20") << QString("[Red]#,##0 ;[Yellow](#,##0)")<<false;
    QTest::newRow("21") << QString("#,##0\\y")<<false;
    QTest::newRow("22") << QString("\"yyyy-mm-dd\"###")<<false;
    QTest::newRow("23") << QString("###;m/d/yy")<<false;
}

QTEST_APPLESS_MAIN(FormatTest)

#include "tst_formattest.moc"
