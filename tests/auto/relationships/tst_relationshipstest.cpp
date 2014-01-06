#include "private/xlsxrelationships_p.h"
#include <QString>
#include <QtTest>

class RelationshipsTest : public QObject
{
    Q_OBJECT

public:
    RelationshipsTest();

private Q_SLOTS:
    void testSaveXml();
    void testLoadXml();
};

RelationshipsTest::RelationshipsTest()
{

}

void RelationshipsTest::testSaveXml()
{
    QXlsx::Relationships rels;
    rels.addDocumentRelationship("/officeDocument", "xl/workbook.xml");

    QByteArray xmldata = rels.saveToXmlData();

    QVERIFY2(xmldata.contains("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"), "");
}

void RelationshipsTest::testLoadXml()
{
    QByteArray xmldata("<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>"
                       "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                       "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                       "</Relationships>");

    QXlsx::Relationships rels;
    rels.loadFromXmlData(xmldata);

    QCOMPARE(rels.documentRelationships("/officeDocument").size(), 1);
}

QTEST_APPLESS_MAIN(RelationshipsTest)

#include "tst_relationshipstest.moc"
