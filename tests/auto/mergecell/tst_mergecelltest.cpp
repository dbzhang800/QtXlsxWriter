#include <QBuffer>
#include <QtTest>

#include "xlsxworksheet.h"
#include "xlsxworkbook.h"

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
    QXlsx::Workbook book;
    QXlsx::Worksheet *sheet = book.addWorksheet("Sheet1");
    sheet->write("B1", "Hello");

    QByteArray xmldata;
    QBuffer buffer(&xmldata);
    buffer.open(QIODevice::WriteOnly);
    sheet->saveToXmlFile(&buffer);

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

void MergeCellTest::testMerge()
{
    QXlsx::Workbook book;
    QXlsx::Worksheet *sheet = book.addWorksheet("Sheet1");
    sheet->write("B1", "Test Merged Cell");
    sheet->mergeCells("B1:B5");

    QByteArray xmldata;
    QBuffer buffer(&xmldata);
    buffer.open(QIODevice::WriteOnly);
    sheet->saveToXmlFile(&buffer);

    QVERIFY2(xmldata.contains("<mergeCells count=\"1\"><mergeCell ref=\"B1:B5\"/></mergeCells>"), "");
}

void MergeCellTest::testUnMerge()
{
    QXlsx::Workbook book;
    QXlsx::Worksheet *sheet = book.addWorksheet("Sheet1");
    sheet->write("B1", "Test Merged Cell");
    sheet->mergeCells("B1:B5");
    sheet->unmergeCells("B1:B5");

    QByteArray xmldata;
    QBuffer buffer(&xmldata);
    buffer.open(QIODevice::WriteOnly);
    sheet->saveToXmlFile(&buffer);

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

QTEST_APPLESS_MAIN(MergeCellTest)

#include "tst_mergecelltest.moc"
