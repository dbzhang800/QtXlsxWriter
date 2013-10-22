#include <QBuffer>
#include <QtTest>

#include "xlsxworksheet.h"
#include "private/xlsxworksheet_p.h"
#include "private/xlsxxmlreader_p.h"

class WorksheetTest : public QObject
{
    Q_OBJECT

public:
    WorksheetTest();

private Q_SLOTS:
    void testEmptySheet();
    void testMerge();
    void testUnMerge();

    void testReadSheetData();
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
            "</row>"
            "</sheetData>";
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is sheetData

    QXlsx::Worksheet sheet("", 0);
    sheet.d_ptr->readSheetData(reader);

    QCOMPARE(sheet.d_ptr->cellTable.size(), 2);
}

QTEST_APPLESS_MAIN(WorksheetTest)

#include "tst_worksheet.moc"
