#include "xlsxdocument.h"
#include <QString>
#include <QtTest>

class ReadDocumentTest : public QObject
{
    Q_OBJECT
    
public:
    ReadDocumentTest();
    
private Q_SLOTS:
    void testDocProps();
};

ReadDocumentTest::ReadDocumentTest()
{
}

void ReadDocumentTest::testDocProps()
{
    QXlsx::Document doc1;
    doc1.setDocumentProperty("creator", "Debao");
    doc1.setDocumentProperty("company", "Test");
    doc1.saveAs("test.xlsx");

    QXlsx::Document doc2("test.xlsx");
    QCOMPARE(doc2.documentProperty("creator"), QString("Debao"));
    QCOMPARE(doc2.documentProperty("company"), QString("Test"));

    QFile::remove("test.xlsx");
}

QTEST_APPLESS_MAIN(ReadDocumentTest)

#include "tst_readdocumenttest.moc"
