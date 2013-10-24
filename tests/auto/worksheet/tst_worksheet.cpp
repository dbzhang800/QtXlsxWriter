#include <QBuffer>
#include <QtTest>

#include "xlsxworksheet.h"
#include "xlsxcell.h"
#include "private/xlsxworksheet_p.h"
#include "private/xlsxxmlreader_p.h"

class WorksheetTest : public QObject
{
    Q_OBJECT

public:
    WorksheetTest();

private Q_SLOTS:
    void testEmptySheet();

    void testWriteCells();
    void testMerge();
    void testUnMerge();

    void testReadSheetData();
    void testReadColsInfo();
    void testReadRowsInfo();
    void testReadMergeCells();
};

WorksheetTest::WorksheetTest()
{
}

void WorksheetTest::testEmptySheet()
{
    QXlsx::Worksheet sheet("", 0);
    sheet.write("B1", 123);
    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

void WorksheetTest::testWriteCells()
{
    QXlsx::Worksheet sheet("", 0);
    sheet.write("A1", 123);
    sheet.write("A2", "Hello");
    sheet.writeInlineString(2, 0, "Hello inline"); //A3
    sheet.write("A4", true);

    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(xmldata.contains("<c r=\"A1\"><v>123</v></c>"), "numeric");
    QVERIFY2(xmldata.contains("<c r=\"A2\" t=\"s\"><v>0</v></c>"), "string");
    QVERIFY2(xmldata.contains("<c r=\"A3\" t=\"inlineStr\"><is><t>Hello inline</t></is></c>"), "inline string");
    QVERIFY2(xmldata.contains("<c r=\"A4\" t=\"b\"><v>1</v></c>"), "boolean");

}

void WorksheetTest::testMerge()
{
    QXlsx::Worksheet sheet("", 0);
    sheet.write("B1", 123);
    sheet.mergeCells("B1:B5");
    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(xmldata.contains("<mergeCells count=\"1\"><mergeCell ref=\"B1:B5\"/></mergeCells>"), "");
}

void WorksheetTest::testUnMerge()
{
    QXlsx::Worksheet sheet("", 0);
    sheet.write("B1", 123);
    sheet.mergeCells("B1:B5");
    sheet.unmergeCells("B1:B5");

    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

void WorksheetTest::testReadSheetData()
{
    const QByteArray xmlData = "<sheetData>"
            "<row r=\"1\" spans=\"1:6\">"
            "<c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c>"
            "</row>"
            "<row r=\"3\" spans=\"1:6\">"
            "<c r=\"B3\" s=\"1\"><v>12345</v></c>"
            "<c r=\"C3\" s=\"1\" t=\"inlineStr\"><is><t>inline test string</t></is></c>"
            "</row>"
            "</sheetData>";
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is sheetData

    QXlsx::Worksheet sheet("", 0);
    sheet.d_ptr->readSheetData(reader);

    QCOMPARE(sheet.d_ptr->cellTable.size(), 2);

    //A1
    QCOMPARE(sheet.d_ptr->cellTable[0][0]->dataType(), QXlsx::Cell::String);
    QCOMPARE(sheet.d_ptr->cellTable[0][0]->value().toInt(), 0);

    //B3
    QCOMPARE(sheet.d_ptr->cellTable[2][1]->dataType(), QXlsx::Cell::Numeric);
    QCOMPARE(sheet.d_ptr->cellTable[2][1]->value().toInt(), 12345);

    //C3
    QCOMPARE(sheet.d_ptr->cellTable[2][2]->dataType(), QXlsx::Cell::InlineString);
    QCOMPARE(sheet.d_ptr->cellTable[2][2]->value().toString(), QStringLiteral("inline test string"));
}

void WorksheetTest::testReadColsInfo()
{
    const QByteArray xmlData = "<cols>"
            "<col min=\"9\" max=\"15\" width=\"5\" style=\"4\" customWidth=\"1\"/>"
            "</cols>";
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is cols

    QXlsx::Worksheet sheet("", 0);
    sheet.d_ptr->readColumnsInfo(reader);

    QCOMPARE(sheet.d_ptr->colsInfo.size(), 1);
    QCOMPARE(sheet.d_ptr->colsInfo[0]->width, 5.0);
}

void WorksheetTest::testReadRowsInfo()
{
    const QByteArray xmlData = "<sheetData>"
            "<row r=\"1\" spans=\"1:6\">"
            "<c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c>"
            "</row>"
            "<row r=\"3\" spans=\"1:6\" s=\"3\" customFormat=\"1\" ht=\"40\" customHeight=\"1\">"
            "<c r=\"B3\" s=\"3\"><v>12345</v></c>"
            "</row>"
            "</sheetData>";
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is sheetData

    QXlsx::Worksheet sheet("", 0);
    sheet.d_ptr->readSheetData(reader);

    QCOMPARE(sheet.d_ptr->rowsInfo.size(), 1);
    QCOMPARE(sheet.d_ptr->rowsInfo[3]->height, 40.0);
}

void WorksheetTest::testReadMergeCells()
{
    const QByteArray xmlData = "<mergeCells count=\"2\"><mergeCell ref=\"B1:B5\"/><mergeCell ref=\"E2:G4\"/></mergeCells>";

    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is mergeCells

    QXlsx::Worksheet sheet("", 0);
    sheet.d_ptr->readMergeCells(reader);

    QCOMPARE(sheet.d_ptr->merges.size(), 2);
    QCOMPARE(sheet.d_ptr->merges[0].row_end, 4);
}

QTEST_APPLESS_MAIN(WorksheetTest)

#include "tst_worksheet.moc"
