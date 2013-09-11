#include <QBuffer>
#include <QtTest>

#include "xlsxworksheet.h"
#include "xlsxdocument.h"

class MergeCellTest : public QObject
{
    Q_OBJECT

public:
    MergeCellTest();

private Q_SLOTS:
    void testWithoutMerge();
    void testMerge();
    void testUnMerge();
};

MergeCellTest::MergeCellTest()
{
}

void MergeCellTest::testWithoutMerge()
{
    QXlsx::Document xlsx;
    xlsx.write("B1", "Hello");

    QByteArray xmldata;
    QBuffer buffer(&xmldata);
    buffer.open(QIODevice::WriteOnly);
    xlsx.activedWorksheet()->saveToXmlFile(&buffer);

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

void MergeCellTest::testMerge()
{
    QXlsx::Document xlsx;
    xlsx.write("B1", "Test Merged Cell");
    xlsx.mergeCells("B1:B5");

    QByteArray xmldata;
    QBuffer buffer(&xmldata);
    buffer.open(QIODevice::WriteOnly);
    xlsx.activedWorksheet()->saveToXmlFile(&buffer);

    QVERIFY2(xmldata.contains("<mergeCells count=\"1\"><mergeCell ref=\"B1:B5\"/></mergeCells>"), "");
}

void MergeCellTest::testUnMerge()
{
    QXlsx::Document xlsx;
    xlsx.write("B1", "Test Merged Cell");
    xlsx.mergeCells("B1:B5");
    xlsx.unmergeCells("B1:B5");

    QByteArray xmldata;
    QBuffer buffer(&xmldata);
    buffer.open(QIODevice::WriteOnly);
    xlsx.activedWorksheet()->saveToXmlFile(&buffer);

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

QTEST_APPLESS_MAIN(MergeCellTest)

#include "tst_mergecelltest.moc"
