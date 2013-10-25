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
    Document xlsx1;
    xlsx1.setDocumentProperty("creator", "Debao");
    xlsx1.setDocumentProperty("company", "Test");
    xlsx1.saveAs("test.xlsx");

    Document xlsx2("test.xlsx");
    QCOMPARE(xlsx2.documentProperty("creator"), QString("Debao"));
    QCOMPARE(xlsx2.documentProperty("company"), QString("Test"));

    QFile::remove("test.xlsx");
}

void DocumentTest::testReadWriteString()
{
    Document xlsx1;
    xlsx1.write("A1", "Hello Qt!");
    Format *format = xlsx1.createFormat();
    format->setFontColor(Qt::blue);
    format->setBorderStyle(Format::BorderDashDotDot);
    format->setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", "Hello Qt again!", format);
    xlsx1.saveAs("test.xlsx");

    Document xlsx2("test.xlsx");
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::String);
    QCOMPARE(xlsx2.cellAt("A1")->value().toString(), QString("Hello Qt!"));
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::String);
    QCOMPARE(xlsx2.cellAt("A2")->value().toString(), QString("Hello Qt again!"));
    QVERIFY(xlsx2.cellAt("A2")->format()!=0);
    QCOMPARE(*xlsx2.cellAt("A2")->format(), *format);

    QFile::remove("test.xlsx");
}

void DocumentTest::testReadWriteNumeric()
{
    Document xlsx1;
    xlsx1.write("A1", 123);
    Format *format = xlsx1.createFormat();
    format->setFontColor(Qt::blue);
    format->setBorderStyle(Format::BorderDashDotDot);
    format->setFillPattern(Format::PatternSolid);
    format->setNumberFormatIndex(10);
    xlsx1.write("A2", 12345, format);
    xlsx1.saveAs("test.xlsx");

    Document xlsx2("test.xlsx");
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Numeric);
    QCOMPARE(xlsx2.cellAt("A1")->value().toDouble(), 123.0);
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Numeric);
    QCOMPARE(xlsx2.cellAt("A2")->value().toDouble(), 12345.0);
    QVERIFY(xlsx2.cellAt("A2")->format()!=0);
    QCOMPARE(*xlsx2.cellAt("A2")->format(), *format);

    QFile::remove("test.xlsx");
}

void DocumentTest::testReadWriteBool()
{
    Document xlsx1;
    xlsx1.write("A1", true);
    Format *format = xlsx1.createFormat();
    format->setFontColor(Qt::blue);
    format->setBorderStyle(Format::BorderDashDotDot);
    format->setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", false, format);
    xlsx1.saveAs("test.xlsx");

    Document xlsx2("test.xlsx");
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Boolean);
    QCOMPARE(xlsx2.cellAt("A1")->value().toBool(), true);
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Boolean);
    QCOMPARE(xlsx2.cellAt("A2")->value().toBool(), false);
    QVERIFY(xlsx2.cellAt("A2")->format()!=0);
    QCOMPARE(*xlsx2.cellAt("A2")->format(), *format);

    QFile::remove("test.xlsx");
}

void DocumentTest::testReadWriteBlank()
{
    Document xlsx1;
    xlsx1.write("A1", QVariant());
    Format *format = xlsx1.createFormat();
    format->setFontColor(Qt::blue);
    format->setBorderStyle(Format::BorderDashDotDot);
    format->setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", QVariant(), format);
    xlsx1.saveAs("test.xlsx");

    Document xlsx2("test.xlsx");
    QVERIFY(xlsx2.cellAt("A1"));
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Blank);
    QVERIFY(!xlsx2.cellAt("A1")->value().isValid());
    QVERIFY(xlsx2.cellAt("A2"));
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Blank);
    QVERIFY(!xlsx2.cellAt("A2")->value().isValid());
    QVERIFY(xlsx2.cellAt("A2")->format()!=0);
    QCOMPARE(*xlsx2.cellAt("A2")->format(), *format);

    QFile::remove("test.xlsx");
}

void DocumentTest::testReadWriteFormula()
{
    Document xlsx1;
    xlsx1.write("A1", "=11+22");
    Format *format = xlsx1.createFormat();
    format->setFontColor(Qt::blue);
    format->setBorderStyle(Format::BorderDashDotDot);
    format->setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", "=22+33", format);
    xlsx1.saveAs("test.xlsx");


    Document xlsx2("test.xlsx");
    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Formula);
//    QCOMPARE(xlsx2.cellAt("A1")->value().toDouble(), 0.0);
    QCOMPARE(xlsx2.cellAt("A1")->formula(), QStringLiteral("11+22"));
    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Formula);
//    QCOMPARE(xlsx2.cellAt("A2")->value().toDouble(), 0.0);
    QCOMPARE(xlsx2.cellAt("A2")->formula(), QStringLiteral("22+33"));
    QVERIFY(xlsx2.cellAt("A2")->format()!=0);
    QCOMPARE(*xlsx2.cellAt("A2")->format(), *format);

    QFile::remove("test.xlsx");
}

void DocumentTest::testReadWriteDateTime()
{
    Document xlsx1;
    QDateTime dt(QDate(2012, 11, 12), QTime(6, 0), Qt::UTC);

    xlsx1.write("A1", dt);

    Format *format = xlsx1.createFormat();
    format->setFontColor(Qt::blue);
    format->setBorderStyle(Format::BorderDashDotDot);
    format->setFillPattern(Format::PatternSolid);
    xlsx1.write("A2", dt, format);

    Format *format3 = xlsx1.createFormat();
    format3->setNumberFormat("dd/mm/yyyy");
    xlsx1.write("A3", dt, format3);

    xlsx1.saveAs("test.xlsx");

    Document xlsx2("test.xlsx");

    QCOMPARE(xlsx2.cellAt("A1")->dataType(), Cell::Numeric);
    QCOMPARE(xlsx2.cellAt("A1")->isDateTime(), true);
    QCOMPARE(xlsx2.cellAt("A1")->dateTime(), dt);

    QCOMPARE(xlsx2.cellAt("A2")->dataType(), Cell::Numeric);
//    QCOMPARE(xlsx2.cellAt("A2")->isDateTime(), true);
//    QCOMPARE(xlsx2.cellAt("A2")->dateTime(), dt);

//    QCOMPARE(xlsx2.cellAt("A3")->dataType(), Cell::Numeric);
//    QVERIFY(xlsx2.cellAt("A3")->format()!=0);
//    qDebug()<<xlsx2.cellAt("A3")->format()->numberFormat();
//    QCOMPARE(xlsx2.cellAt("A3")->isDateTime(), true);
//    QCOMPARE(xlsx2.cellAt("A3")->dateTime(), dt);
//    QCOMPARE(xlsx2.cellAt("A3")->format()->numberFormat(), QString("dd/mm/yyyy"));

    QFile::remove("test.xlsx");

}

QTEST_APPLESS_MAIN(DocumentTest)

#include "tst_documenttest.moc"
