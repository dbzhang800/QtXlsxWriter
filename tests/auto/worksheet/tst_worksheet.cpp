#include <QBuffer>
#include <QtTest>
#include <QXmlStreamReader>

#include "xlsxworksheet.h"
#include "xlsxcell.h"
#include "xlsxcellrange.h"
#include "xlsxdatavalidation.h"
#include "private/xlsxworksheet_p.h"
#include "private/xlsxsharedstrings_p.h"
#include "xlsxrichstring.h"
#include "xlsxcellformula.h"

class WorksheetTest : public QObject
{
    Q_OBJECT

public:
    WorksheetTest();

private Q_SLOTS:
    void testEmptySheet();
    void testDimension();
    void testSheetView();
    void testSetColumn();

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
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
    sheet.write("B1", 123);
    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(!xmldata.contains("<mergeCell"), "");
}

void WorksheetTest::testDimension()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
    QCOMPARE(sheet.dimension(), QXlsx::CellRange()); //Default

    sheet.write("C3", "Test");
    qDebug()<<sheet.dimension().toString();
    QCOMPARE(sheet.dimension(), QXlsx::CellRange(3, 3, 3, 3)); //Single Cell

    sheet.write("B2", "Second");
    QCOMPARE(sheet.dimension(), QXlsx::CellRange(2, 2, 3, 3));

    sheet.write("D4", "Test");
    QCOMPARE(sheet.dimension(), QXlsx::CellRange("B2:D4"));

    sheet.write(10000, 10000, "For test");
    QCOMPARE(sheet.dimension(), QXlsx::CellRange(2, 2, 10000, 10000));
}

void WorksheetTest::testSheetView()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
    sheet.setGridLinesVisible(false);
    sheet.setWindowProtected(true);
    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(xmldata.contains("showGridLines=\"0\""), "gridlines");
    QVERIFY2(xmldata.contains("windowProtection=\"1\""), "windowProtection");
}

void WorksheetTest::testSetColumn()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
    sheet.setColumnWidth(1, 11, 20.0); //"A:K"
    sheet.setColumnWidth(4, 8, 10.0); //"D:H"
    sheet.setColumnWidth(6, 6, 15.0); //"F:F"
    sheet.setColumnWidth(1, 9, 8.8); //"A:H"

    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY(xmldata.contains("<col min=\"1\" max=\"3\"")); //"A:C"
    QVERIFY(xmldata.contains("<col min=\"4\" max=\"5\"")); //"D:E"
    QVERIFY(xmldata.contains("<col min=\"6\" max=\"6\"")); //"F:F"
    QVERIFY(xmldata.contains("<col min=\"7\" max=\"8\"")); //"G:H"
    QVERIFY(xmldata.contains("<col min=\"9\" max=\"9\""));//"I:I"
    QVERIFY(xmldata.contains("<col min=\"10\" max=\"11\""));//"J:K"
}

void WorksheetTest::testWriteCells()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
    sheet.write("A1", 123);
    sheet.write("A2", "Hello");
    sheet.writeInlineString(3, 1, "Hello inline"); //A3
    sheet.write("A4", true);
    sheet.write("A5", "=44+33");
    sheet.writeFormula(5, 2, "44+33", QXlsx::Format(), 77);

    QByteArray xmldata = sheet.saveToXmlData();
    qDebug()<<xmldata;

    QVERIFY2(xmldata.contains("<c r=\"A1\"><v>123</v></c>"), "numeric");
    QVERIFY2(xmldata.contains("<c r=\"A2\" t=\"s\"><v>0</v></c>"), "string");
    QVERIFY2(xmldata.contains("<c r=\"A3\" t=\"inlineStr\"><is><t>Hello inline</t></is></c>"), "inline string");
    QVERIFY2(xmldata.contains("<c r=\"A4\" t=\"b\"><v>1</v></c>"), "boolean");
    QVERIFY2(xmldata.contains("<c r=\"A5\"><f ca=\"1\">44+33</f><v>0</v></c>"), "formula");
    QVERIFY2(xmldata.contains("<c r=\"B5\"><f ca=\"1\">44+33</f><v>77</v></c>"), "formula");

    QCOMPARE(sheet.d_func()->sharedStrings()->getSharedString(0).toPlainString(), QStringLiteral("Hello"));
}

void WorksheetTest::testWriteHyperlinks()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
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

    QCOMPARE(sheet.d_func()->sharedStrings()->getSharedString(0).toPlainString(), QStringLiteral("http://qt-project.org"));
    QCOMPARE(sheet.d_func()->sharedStrings()->getSharedString(1).toPlainString(), QStringLiteral("http://qt-project.org/abc"));
    QCOMPARE(sheet.d_func()->sharedStrings()->getSharedString(2).toPlainString(), QStringLiteral("http://qt-project.org/abc.html#test"));
    QCOMPARE(sheet.d_func()->sharedStrings()->getSharedString(3).toPlainString(), QStringLiteral("xyz@debao.me"));
    QCOMPARE(sheet.d_func()->sharedStrings()->getSharedString(4).toPlainString(), QStringLiteral("xyz@debao.me?subject=Test"));
}

void WorksheetTest::testWriteDataValidations()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
    QXlsx::DataValidation validation(QXlsx::DataValidation::Whole);
    validation.setFormula1("10");
    validation.setFormula2("100");
    validation.addCell("A1");
    validation.addRange("C2:C4");
    sheet.addDataValidation(validation);

    QByteArray xmldata = sheet.saveToXmlData();
    QVERIFY(xmldata.contains("<dataValidation type=\"whole\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1 C2:C4\"><formula1>10</formula1><formula2>100</formula2></dataValidation>"));
 }

void WorksheetTest::testMerge()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
    sheet.write("B1", 123);
    sheet.mergeCells("B1:B5");
    QByteArray xmldata = sheet.saveToXmlData();

    QVERIFY2(xmldata.contains("<mergeCells count=\"1\"><mergeCell ref=\"B1:B5\"/></mergeCells>"), "");
}

void WorksheetTest::testUnMerge()
{
    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_NewFromScratch);
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
    QXmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is sheetData

    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_LoadFromExists);
    sheet.d_func()->sharedStrings()->addSharedString("Hello");
    sheet.d_func()->loadXmlSheetData(reader);

    QCOMPARE(sheet.d_func()->cellTable.size(), 2);

    //A1
    QCOMPARE(sheet.cellAt("A1")->cellType(), QXlsx::Cell::SharedStringType);
    QCOMPARE(sheet.cellAt("A1")->value().toString(), QStringLiteral("Hello"));

    //B1
    QCOMPARE(sheet.cellAt("B1")->cellType(), QXlsx::Cell::NumberType);
    QCOMPARE(sheet.cellAt("B1")->value().toInt(), 77);
    QCOMPARE(sheet.cellAt("B1")->formula(), QXlsx::CellFormula("44+33"));

    //C1
    QCOMPARE(sheet.cellAt("C1")->cellType(), QXlsx::Cell::StringType);
    QCOMPARE(sheet.cellAt("C1")->value().toInt(), 77);
    QCOMPARE(sheet.cellAt("C1")->formula(), QXlsx::CellFormula("44+33"));

    //B3
    QCOMPARE(sheet.cellAt("B3")->cellType(), QXlsx::Cell::NumberType);
    QCOMPARE(sheet.cellAt("B3")->value().toInt(), 12345);

    //C3
    QCOMPARE(sheet.cellAt("C3")->cellType(), QXlsx::Cell::InlineStringType);
    QCOMPARE(sheet.cellAt("C3")->value().toString(), QStringLiteral("inline test string"));

    //E3
    QCOMPARE(sheet.cellAt("E3")->cellType(), QXlsx::Cell::ErrorType);
    QCOMPARE(sheet.cellAt("E3")->value().toString(), QStringLiteral("#DIV/0!"));
}

void WorksheetTest::testReadColsInfo()
{
    const QByteArray xmlData = "<cols>"
            "<col min=\"9\" max=\"15\" width=\"5\" style=\"4\" customWidth=\"1\"/>"
            "</cols>";
    QXmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is cols

    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_LoadFromExists);
    sheet.d_func()->loadXmlColumnsInfo(reader);

    QCOMPARE(sheet.d_func()->colsInfo.size(), 1);
    QCOMPARE(sheet.d_func()->colsInfo[9]->width, 5.0);
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
    QXmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is sheetData

    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_LoadFromExists);
    sheet.d_func()->loadXmlSheetData(reader);

    QCOMPARE(sheet.d_func()->rowsInfo.size(), 1);
    QCOMPARE(sheet.d_func()->rowsInfo[3]->height, 40.0);
}

void WorksheetTest::testReadMergeCells()
{
    const QByteArray xmlData = "<mergeCells count=\"2\"><mergeCell ref=\"B1:B5\"/><mergeCell ref=\"E2:G4\"/></mergeCells>";

    QXmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is mergeCells

    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_LoadFromExists);
    sheet.d_func()->loadXmlMergeCells(reader);

    QCOMPARE(sheet.d_func()->merges.size(), 2);
    QCOMPARE(sheet.d_func()->merges[0].toString(), QStringLiteral("B1:B5"));
}

void WorksheetTest::testReadDataValidations()
{
    const QByteArray xmlData = "<dataValidations count=\"2\">"
            "<dataValidation type=\"whole\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1 C2:C4\"><formula1>10</formula1><formula2>100</formula2></dataValidation>"
            "<dataValidation type=\"whole\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1 C2:C4\"><formula1>10</formula1><formula2>100</formula2></dataValidation>"
            "</dataValidations>";

    QXmlStreamReader reader(xmlData);
    reader.readNextStartElement();//current node is dataValidations

    QXlsx::Worksheet sheet("", 1, 0, QXlsx::Worksheet::F_LoadFromExists);
    sheet.d_func()->loadXmlDataValidations(reader);

    QCOMPARE(sheet.d_func()->dataValidationsList.size(), 2);
    QCOMPARE(sheet.d_func()->dataValidationsList[0].validationType(), QXlsx::DataValidation::Whole);
}

QTEST_APPLESS_MAIN(WorksheetTest)

#include "tst_worksheet.moc"
