#include "xlsxdocument.h"
#include "xlsxcell.h"
#include "xlsxformat.h"
#include "xlsxcellformula.h"
#include <QString>
#include <QtTest>

QTXLSX_USE_NAMESPACE

class DocumentTest : public QObject
{
    Q_OBJECT
    
public:
    DocumentTest();
    
private Q_SLOTS:
    void testDocumentProperty();
    void testReadWriteString();
    void testReadWriteNumeric();
    void testReadWriteBool();
    void testReadWriteBlank();
    void testReadWriteFormula();
    void testReadWriteDateTime();
    void testReadWriteDate();
    void testReadWriteTime();

    void testMoveWorksheet();
    void testDeleteWorksheet();
    void testCopyWorksheet();
};

DocumentTest::DocumentTest()
{
}

void DocumentTest::testDocumentProperty()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    xlsx1.setDocumentProperty("creator", "Debao");
    xlsx1.setDocumentProperty("company", "Test");
    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.documentProperty("creator"), QString("Debao"));
    QCOMPARE(xlsx2.documentProperty("company"), QString("Test"));
}

void DocumentTest::testReadWriteString()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    xlsx1.write("A1", "Hello Qt!");

    Format format;
    format.setFontColor(Qt::blue);
    format.setBorderStyle(Format::BorderDashDotDot);
    format.setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", "Hello Qt again!", format);

    xlsx1.write("A3", "12345");

    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::SharedStringType);
    QCOMPARE(xlsx2.cellAt("A1")->value().toString(), QString("Hello Qt!"));
    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::SharedStringType);
    QCOMPARE(xlsx2.cellAt("A2")->value().toString(), QString("Hello Qt again!"));
    Format format2 = xlsx2.cellAt("A2")->format();
    QVERIFY(format2.isValid());
//    qDebug()<<format2;
//    qDebug()<<format;
    QCOMPARE(format2, format);

    QCOMPARE(xlsx2.cellAt("A3")->cellType(), Cell::SharedStringType);
    QCOMPARE(xlsx2.cellAt("A3")->value().toString(), QString("12345"));
}

void DocumentTest::testReadWriteNumeric()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    xlsx1.write("A1", 123);
    Format format;
    format.setFontColor(Qt::blue);
    format.setBorderStyle(Format::BorderDashDotDot);
    format.setFillPattern(Format::PatternSolid);
    format.setNumberFormatIndex(10);
    xlsx1.write("A2", 12345, format);
    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A1")->value().toDouble(), 123.0);
    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A2")->value().toDouble(), 12345.0);
    QVERIFY(xlsx2.cellAt("A2")->format().isValid());
    QCOMPARE(xlsx2.cellAt("A2")->format(), format);
}

void DocumentTest::testReadWriteBool()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    xlsx1.write("A1", true);
    Format format;
    format.setFontColor(Qt::blue);
    format.setBorderStyle(Format::BorderDashDotDot);
    format.setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", false, format);
    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::BooleanType);
    QCOMPARE(xlsx2.cellAt("A1")->value().toBool(), true);
    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::BooleanType);
    QCOMPARE(xlsx2.cellAt("A2")->value().toBool(), false);
    QVERIFY(xlsx2.cellAt("A2")->format().isValid());
    QCOMPARE(xlsx2.cellAt("A2")->format(), format);
}

void DocumentTest::testReadWriteBlank()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    xlsx1.write("A1", QVariant());
    Format format;
    format.setFontColor(Qt::blue);
    format.setBorderStyle(Format::BorderDashDotDot);
    format.setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", QVariant(), format);
    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QVERIFY(xlsx2.cellAt("A1"));
    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::NumberType);
    QVERIFY(!xlsx2.cellAt("A1")->value().isValid());
    QVERIFY(xlsx2.cellAt("A2"));
    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::NumberType);
    QVERIFY(!xlsx2.cellAt("A2")->value().isValid());
    QVERIFY(xlsx2.cellAt("A2")->format().isValid());
    QCOMPARE(xlsx2.cellAt("A2")->format(), format);
}

void DocumentTest::testReadWriteFormula()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    xlsx1.write("A1", "=11+22");
    Format format;
    format.setFontColor(Qt::blue);
    format.setBorderStyle(Format::BorderDashDotDot);
    format.setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", "=22+33", format);
    xlsx1.saveAs(&device);


    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::NumberType);
    QVERIFY(xlsx2.cellAt("A1")->hasFormula());
    QCOMPARE(xlsx2.cellAt("A1")->formula(), CellFormula("11+22"));
    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A2")->formula(), CellFormula("22+33"));
    QVERIFY(xlsx2.cellAt("A2")->format().isValid());
    QCOMPARE(xlsx2.cellAt("A2")->format(), format);
}

void DocumentTest::testReadWriteDateTime()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    QDateTime dt(QDate(2012, 11, 12), QTime(6, 0));

    xlsx1.write("A1", dt);

    Format format;
    format.setFontColor(Qt::blue);
    format.setBorderStyle(Format::BorderDashDotDot);
    format.setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", dt, format);

    Format format3;
    format3.setNumberFormat("dd/mm/yyyy");
    xlsx1.write("A3", dt, format3);

//    xlsx1.write("A4", "2013-12-14T12:30"); //Auto convert to QDateTime, by QVariant

    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A1")->isDateTime(), true);
    QCOMPARE(xlsx2.cellAt("A1")->dateTime(), dt);
    QVERIFY(xlsx2.read("A1").userType() == QMetaType::QDateTime);

    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A2")->isDateTime(), true);
    QCOMPARE(xlsx2.cellAt("A2")->dateTime(), dt);
    QVERIFY(xlsx2.read("A2").userType() == QMetaType::QDateTime);

    QCOMPARE(xlsx2.cellAt("A3")->cellType(), Cell::NumberType);
    QVERIFY(xlsx2.cellAt("A3")->format().isValid());
    QCOMPARE(xlsx2.cellAt("A3")->isDateTime(), true);
    QCOMPARE(xlsx2.cellAt("A3")->dateTime(), dt);
    QCOMPARE(xlsx2.cellAt("A3")->format().numberFormat(), QString("dd/mm/yyyy"));

//    QCOMPARE(xlsx2.cellAt("A4")->dataType(), Cell::Numeric);
//    QCOMPARE(xlsx2.cellAt("A4")->isDateTime(), true);
//    QCOMPARE(xlsx2.cellAt("A4")->dateTime(), QDateTime(QDate(2013,12,14), QTime(12, 30)));
//    QVERIFY(xlsx2.read("A4").userType() == QMetaType::QDateTime);
}

void DocumentTest::testReadWriteDate()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;
    QDate d(2012, 11, 12);

    xlsx1.write("A1", d);

    Format format;
    format.setFontColor(Qt::blue);
    format.setBorderStyle(Format::BorderDashDotDot);
    format.setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", d, format);

    Format format3;
    format3.setNumberFormat("dd/mm/yyyy");
    xlsx1.write("A3", d, format3);

//    xlsx1.write("A4", "2013-12-14"); //Auto convert to QDateTime, by QVariant

    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A1")->isDateTime(), true);
    QVERIFY(xlsx2.read("A1").userType() == QMetaType::QDate);
    QCOMPARE(xlsx2.read("A1").toDate(), d);

    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A2")->isDateTime(), true);
    QVERIFY(xlsx2.read("A2").userType() == QMetaType::QDate);

    QCOMPARE(xlsx2.cellAt("A3")->cellType(), Cell::NumberType);
    QVERIFY(xlsx2.cellAt("A3")->format().isValid());
    QCOMPARE(xlsx2.cellAt("A3")->isDateTime(), true);
    QCOMPARE(xlsx2.cellAt("A3")->format().numberFormat(), QString("dd/mm/yyyy"));
    QVERIFY(xlsx2.read("A3").userType() == QMetaType::QDate);
    QCOMPARE(xlsx2.read("A3").toDate(), d);

//    QCOMPARE(xlsx2.cellAt("A4")->dataType(), Cell::Numeric);
//    QCOMPARE(xlsx2.cellAt("A4")->isDateTime(), true);
//    QCOMPARE(xlsx2.cellAt("A4")->dateTime(), QDateTime(QDate(2013,12,14)));
//    QVERIFY(xlsx2.read("A4").userType() == QMetaType::QDate);
//    QCOMPARE(xlsx2.read("A4").toDate(), QDate(2013,12,14));
}

void DocumentTest::testReadWriteTime()
{
    QBuffer device;
    device.open(QIODevice::WriteOnly);

    Document xlsx1;

    xlsx1.write("A1", QTime()); //Blank cell
    xlsx1.write("A2", QTime(1, 22));

    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);

    QCOMPARE(xlsx2.cellAt("A1")->cellType(), Cell::NumberType);
    QVERIFY(!xlsx2.cellAt("A1")->value().isValid());

    QCOMPARE(xlsx2.cellAt("A2")->cellType(), Cell::NumberType);
    QCOMPARE(xlsx2.cellAt("A2")->isDateTime(), true);
    QVERIFY(xlsx2.read("A2").userType() == QMetaType::QTime);
    QCOMPARE(xlsx2.read("A2").toTime(), QTime(1, 22));
}

void DocumentTest::testMoveWorksheet()
{
    Document xlsx1;
    xlsx1.addSheet();//Sheet1
    xlsx1.addSheet();

    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet1"<<"Sheet2");
    xlsx1.moveSheet("Sheet2", 0);
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet2"<<"Sheet1");
    xlsx1.moveSheet("Sheet2", 1);
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet1"<<"Sheet2");
}

void DocumentTest::testCopyWorksheet()
{
    Document xlsx1;
    xlsx1.addSheet();//Sheet1
    xlsx1.addSheet();
    xlsx1.write("A1", "String");
    xlsx1.write("A2", 999);
    xlsx1.write("A3", true);
    xlsx1.addSheet();
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet1"<<"Sheet2"<<"Sheet3");

    xlsx1.copySheet("Sheet2");
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet1"<<"Sheet2"<<"Sheet3"<<"Sheet2(2)");

    xlsx1.deleteSheet("Sheet2");
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet1"<<"Sheet3"<<"Sheet2(2)");

    xlsx1.selectSheet("Sheet2(2)");
    QCOMPARE(xlsx1.read("A1").toString(), QString("String"));
    QCOMPARE(xlsx1.read("A2").toInt(), 999);
    QCOMPARE(xlsx1.read("A3").toBool(), true);
}

void DocumentTest::testDeleteWorksheet()
{
    Document xlsx1;
    xlsx1.addSheet();//Sheet1
    xlsx1.addSheet();
    xlsx1.addSheet();

    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet1"<<"Sheet2"<<"Sheet3");
    xlsx1.deleteSheet("Sheet2");
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet1"<<"Sheet3");
    xlsx1.deleteSheet("Sheet1");
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet3");

    //Cann't delete the last worksheet
    xlsx1.deleteSheet("Sheet3");
    QCOMPARE(xlsx1.sheetNames(), QStringList()<<"Sheet3");
}

QTEST_APPLESS_MAIN(DocumentTest)

#include "tst_documenttest.moc"
