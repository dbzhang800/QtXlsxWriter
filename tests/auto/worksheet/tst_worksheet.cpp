#include <QBuffer>
#include <QtTest>

#include "xlsxworksheet.h"
#include "xlsxcell.h"
#include "xlsxdatavalidation.h"
#include "private/xlsxworksheet_p.h"
#include "private/xlsxxmlreader_p.h"
#include "private/xlsxsharedstrings_p.h"

class WorksheetTest : public QObject
{
    Q_OBJECT

public:
    WorksheetTest();

private Q_SLOTS:
    void testEmptySheet();

    void testFirstRow();
    void testLastRow();
    void testFirstColumn();
    void testLastColumn();

    void testWriteCells();
    void testWriteHyperlinks();
    void testWriteDataValidations();
    void testMerge();
    void testUnMerge();

    void testReadSheetData();
    void testReadColsInfo();
    void testReadRowsInfo();
    void testReadMergeCells();
    void testReadDataValidations();
};

WorksheetTest::WorksheetTest()
{
}

void WorksheetTest::testEmptySheet()
{
    QXlsx::Worksheet sheet("", 1, 0);
    sheet.write("B1", 123);
    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

void WorksheetTest::testFirstRow()
{
    QXlsx::Worksheet sheet("", 1, 0);
    QCOMPARE(sheet.firstRow(), 0); //Default

    sheet.write(10000, 10000, "For test");
    QCOMPARE(sheet.firstRow(), 10000);

    sheet.write("C3", "Test");
    QCOMPARE(sheet.firstRow(), 2); //Single Cell

    sheet.write("B2", "Second");
    QCOMPARE(sheet.firstRow(), 1);

    sheet.write("D4", "Test");
    QCOMPARE(sheet.firstRow(), 1);
}

void WorksheetTest::testLastRow()
{
    QXlsx::Worksheet sheet("", 1, 0);
    QCOMPARE(sheet.lastRow(), 0); //Default

    sheet.write("C3", "Test");
    QCOMPARE(sheet.lastRow(), 3); //Single Cell

    sheet.write("B2", "Second");
    QCOMPARE(sheet.lastRow(), 3);

    sheet.write("D4", "Test");
    QCOMPARE(sheet.lastRow(), 4);

    sheet.write(10000, 10000, "For test");
    QCOMPARE(sheet.lastRow(), 10001);
}

void WorksheetTest::testFirstColumn()
{
    QXlsx::Worksheet sheet("", 1, 0);
    QCOMPARE(sheet.firstColumn(), 0); //Default

    sheet.write(10000, 10000, "For test");
    QCOMPARE(sheet.firstColumn(), 10000);

    sheet.write("C3", "Test");
    QCOMPARE(sheet.firstColumn(), 2); //Single Cell

    sheet.write("B2", "Second");
    QCOMPARE(sheet.firstColumn(), 1);

    sheet.write("D4", "Test");
    QCOMPARE(sheet.firstColumn(), 1);
}

void WorksheetTest::testLastColumn()
{
    QXlsx::Worksheet sheet("", 1, 0);
    QCOMPARE(sheet.lastColumn(), 0); //Default

    sheet.write("C3", "Test");
    QCOMPARE(sheet.lastColumn(), 3); //Single Cell

    sheet.write("B2", "Second");
    QCOMPARE(sheet.lastColumn(), 3);

    sheet.write("D4", "Test");
    QCOMPARE(sheet.lastColumn(), 4);

    sheet.write(10000, 10000, "For test");
    QCOMPARE(sheet.lastColumn(), 10001);
}

void WorksheetTest::testWriteCells()
{
    QXlsx::Worksheet sheet("", 1, 0);
    sheet.write("A1", 123);
    sheet.write("A2", "Hello");
    sheet.writeInlineString(2, 0, "Hello inline"); //A3
    sheet.write("A4", true);
    sheet.write("A5", "=44+33");
    sheet.writeFormula(4, 1, "44+33", 0, 77);

    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(xmldata.contains("<c r=\"A1\"><v>123</v></c>"), "numeric");
    QVERIFY2(xmldata.contains("<c r=\"A2\" t=\"s\"><v>0</v></c>"), "string");
    QVERIFY2(xmldata.contains("<c r=\"A3\" t=\"inlineStr\"><is><t>Hello inline</t></is></c>"), "inline string");
    QVERIFY2(xmldata.contains("<c r=\"A4\" t=\"b\"><v>1</v></c>"), "boolean");
    QVERIFY2(xmldata.contains("<c r=\"A5\" t=\"str\"><f>44+33</f><v>0</v></c>"), "formula");
    QVERIFY2(xmldata.contains("<c r=\"B5\" t=\"str\"><f>44+33</f><v>77</v></c>"), "formula");

    QCOMPARE(sheet.d_ptr->sharedStrings()->getSharedString(0), QStringLiteral("Hello"));
}

void WorksheetTest::testWriteHyperlinks()
{
    QXlsx::Worksheet sheet("", 1, 0);
    sheet.write("A1", QUrl::fromUserInput("http://qt-project.org"));
    sheet.write("B1", QUrl::fromUserInput("http://qt-project.org/abc"));
    sheet.write("C1", QUrl::fromUserInput("http://qt-project.org/abc.html#test"));
    sheet.write("D1", QUrl::fromUserInput("mailto:xyz@debao.me"));
    sheet.write("E1", QUrl::fromUserInput("mailto:xyz@debao.me?subject=Test"));

    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(xmldata.contains("<hyperlink ref=\"A1\" r:id=\"rId1\"/>"), "simple");
    QVERIFY2(xmldata.contains("<hyperlink ref=\"B1\" r:id=\"rId2\"/>"), "url with path");
    QVERIFY2(xmldata.contains("<hyperlink ref=\"C1\" r:id=\"rId3\" location=\"test\"/>"), "url with location");
    QVERIFY2(xmldata.contains("<hyperlink ref=\"D1\" r:id=\"rId4\"/>"), "mail");
    QVERIFY2(xmldata.contains("<hyperlink ref=\"E1\" r:id=\"rId5\"/>"), "mail with subject");

    QCOMPARE(sheet.d_ptr->sharedStrings()->getSharedString(0), QStringLiteral("http://qt-project.org"));
    QCOMPARE(sheet.d_ptr->sharedStrings()->getSharedString(1), QStringLiteral("http://qt-project.org/abc"));
    QCOMPARE(sheet.d_ptr->sharedStrings()->getSharedString(2), QStringLiteral("http://qt-project.org/abc.html#test"));
    QCOMPARE(sheet.d_ptr->sharedStrings()->getSharedString(3), QStringLiteral("xyz@debao.me"));
    QCOMPARE(sheet.d_ptr->sharedStrings()->getSharedString(4), QStringLiteral("xyz@debao.me?subject=Test"));
}

void WorksheetTest::testWriteDataValidations()
{
    QXlsx::Worksheet sheet("", 1, 0);
    QXlsx::DataValidation validation(QXlsx::DataValidation::Whole);
    validation.setFormula1("10");
    validation.setFormula2("100");
    validation.addCell("A1");
    validation.addCell("C2:C4");
    sheet.addDataValidation(validation);

    QByteArray xmldata = sheet.saveToXmlData();
    QVERIFY(xmldata.contains("<dataValidation type=\"whole\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1 C2:C4\"><formula1>10</formula1><formula2>100</formula2></dataValidation>"));
 }

void WorksheetTest::testMerge()
{
    QXlsx::Worksheet sheet("", 1, 0);
    sheet.write("B1", 123);
    sheet.mergeCells("B1:B5");
    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(xmldata.contains("<mergeCells count=\"1\"><mergeCell ref=\"B1:B5\"/></mergeCells>"), "");
}

void WorksheetTest::testUnMerge()
{
    QXlsx::Worksheet sheet("", 1, 0);
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
            "<c r=\"B1\"><f>44+33</f><v>77</v></c>"
            "<c r=\"C1\" t=\"str\"><f>44+33</f><v>77</v></c>"
            "</row>"
            "<row r=\"3\" spans=\"1:6\">"
            "<c r=\"B3\" s=\"1\"><v>12345</v></c>"
            "<c r=\"C3\" s=\"1\" t=\"inlineStr\"><is><t>inline test string</t></is></c>"
            "<c r=\"E3\" t=\"e\"><f>1/0</f><v>#DIV/0!</v></c>"
            "</row>"
            "</sheetData>";
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is sheetData

    QXlsx::Worksheet sheet("", 1, 0);
    sheet.d_ptr->sharedStrings()->addSharedString("Hello");
    sheet.d_ptr->readSheetData(reader);

    QCOMPARE(sheet.d_ptr->cellTable.size(), 2);

    //A1
    QCOMPARE(sheet.cellAt("A1")->dataType(), QXlsx::Cell::String);
    QCOMPARE(sheet.cellAt("A1")->value().toString(), QStringLiteral("Hello"));

    //B1
    QCOMPARE(sheet.cellAt("B1")->dataType(), QXlsx::Cell::Formula);
    QCOMPARE(sheet.cellAt("B1")->value().toInt(), 77);
    QCOMPARE(sheet.cellAt("B1")->formula(), QStringLiteral("44+33"));

    //C1
    QCOMPARE(sheet.cellAt("C1")->dataType(), QXlsx::Cell::Formula);
    QCOMPARE(sheet.cellAt("C1")->value().toInt(), 77);
    QCOMPARE(sheet.cellAt("C1")->formula(), QStringLiteral("44+33"));

    //B3
    QCOMPARE(sheet.cellAt("B3")->dataType(), QXlsx::Cell::Numeric);
    QCOMPARE(sheet.cellAt("B3")->value().toInt(), 12345);

    //C3
    QCOMPARE(sheet.cellAt("C3")->dataType(), QXlsx::Cell::InlineString);
    QCOMPARE(sheet.cellAt("C3")->value().toString(), QStringLiteral("inline test string"));

    //E3
    QCOMPARE(sheet.cellAt("E3")->dataType(), QXlsx::Cell::Error);
    QCOMPARE(sheet.cellAt("E3")->value().toString(), QStringLiteral("#DIV/0!"));
}

void WorksheetTest::testReadColsInfo()
{
    const QByteArray xmlData = "<cols>"
            "<col min=\"9\" max=\"15\" width=\"5\" style=\"4\" customWidth=\"1\"/>"
            "</cols>";
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is cols

    QXlsx::Worksheet sheet("", 1, 0);
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

    QXlsx::Worksheet sheet("", 1, 0);
    sheet.d_ptr->readSheetData(reader);

    QCOMPARE(sheet.d_ptr->rowsInfo.size(), 1);
    QCOMPARE(sheet.d_ptr->rowsInfo[3]->height, 40.0);
}

void WorksheetTest::testReadMergeCells()
{
    const QByteArray xmlData = "<mergeCells count=\"2\"><mergeCell ref=\"B1:B5\"/><mergeCell ref=\"E2:G4\"/></mergeCells>";

    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is mergeCells

    QXlsx::Worksheet sheet("", 1, 0);
    sheet.d_ptr->readMergeCells(reader);

    QCOMPARE(sheet.d_ptr->merges.size(), 2);
    QCOMPARE(sheet.d_ptr->merges[0].toString(), QStringLiteral("B1:B5"));
}

void WorksheetTest::testReadDataValidations()
{
    const QByteArray xmlData = "<dataValidations count=\"2\">"
            "<dataValidation type=\"whole\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1 C2:C4\"><formula1>10</formula1><formula2>100</formula2></dataValidation>"
            "<dataValidation type=\"whole\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1 C2:C4\"><formula1>10</formula1><formula2>100</formula2></dataValidation>"
            "</dataValidations>";

    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is dataValidations

    QXlsx::Worksheet sheet("", 1, 0);
    sheet.d_ptr->readDataValidations(reader);

    QCOMPARE(sheet.d_ptr->dataValidationsList.size(), 2);
    QCOMPARE(sheet.d_ptr->dataValidationsList[0].validationType(), QXlsx::DataValidation::Whole);
}

QTEST_APPLESS_MAIN(WorksheetTest)

#include "tst_worksheet.moc"
