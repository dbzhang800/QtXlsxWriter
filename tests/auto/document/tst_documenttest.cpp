#include "xlsxdocument.h"
#include "xlsxcell.h"
#include "xlsxformat.h"
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
    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::String);
    QCOMPARE(xlsx2.cellAt("A1")->value().toString(), QString("Hello Qt!"));
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::String);
    QCOMPARE(xlsx2.cellAt("A2")->value().toString(), QString("Hello Qt again!"));
    Format format2 = xlsx2.cellAt("A2")->format();
    QVERIFY(format2.isValid());
//    qDebug()<<format2;
//    qDebug()<<format;
    QCOMPARE(format2, format);
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
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Numeric);
    QCOMPARE(xlsx2.cellAt("A1")->value().toDouble(), 123.0);
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Numeric);
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
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Boolean);
    QCOMPARE(xlsx2.cellAt("A1")->value().toBool(), true);
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Boolean);
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
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Blank);
    QVERIFY(!xlsx2.cellAt("A1")->value().isValid());
    QVERIFY(xlsx2.cellAt("A2"));
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Blank);
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
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Formula);
//    QCOMPARE(xlsx2.cellAt("A1")->value().toDouble(), 0.0);
    QCOMPARE(xlsx2.cellAt("A1")->formula(), QStringLiteral("11+22"));
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Formula);
//    QCOMPARE(xlsx2.cellAt("A2")->value().toDouble(), 0.0);
    QCOMPARE(xlsx2.cellAt("A2")->formula(), QStringLiteral("22+33"));
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

    xlsx1.saveAs(&device);

    device.open(QIODevice::ReadOnly);
    Document xlsx2(&device);
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Numeric);
    QCOMPARE(xlsx2.cellAt("A1")->isDateTime(), true);
    QCOMPARE(xlsx2.cellAt("A1")->dateTime(), dt);

    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Numeric);
//    QCOMPARE(xlsx2.cellAt("A2")->isDateTime(), true);
//    QCOMPARE(xlsx2.cellAt("A2")->dateTime(), dt);

    QCOMPARE(xlsx2.cellAt("A3")->dataType(), Cell::Numeric);
    QVERIFY(xlsx2.cellAt("A3")->format().isValid());
    qDebug()<<xlsx2.cellAt("A3")->format().numberFormat();
    QCOMPARE(xlsx2.cellAt("A3")->isDateTime(), true);
    QCOMPARE(xlsx2.cellAt("A3")->dateTime(), dt);
    QCOMPARE(xlsx2.cellAt("A3")->format().numberFormat(), QString("dd/mm/yyyy"));
}

QTEST_APPLESS_MAIN(DocumentTest)

#include "tst_documenttest.moc"
