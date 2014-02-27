#include "private/xlsxsharedstrings_p.h"
#include "xlsxrichstring.h"
#include "xlsxformat.h"
#include <QString>
#include <QtTest>
#include <QXmlStreamReader>

class SharedStringsTest : public QObject
{
    Q_OBJECT

public:
    SharedStringsTest();

private Q_SLOTS:
    void testAddSharedString();
    void testRemoveSharedString();

    void testLoadXmlData();
    void testLoadRichStringXmlData();

};

SharedStringsTest::SharedStringsTest()
{
}

void SharedStringsTest::testAddSharedString()
{
    QXlsx::SharedStrings sst(QXlsx::SharedStrings::F_NewFromScratch);
    sst.addSharedString("Hello Qt!");
    sst.addSharedString("Xlsx Writer");

    QXlsx::RichString rs;
    rs.addFragment("Hello", QXlsx::Format());
    rs.addFragment(" RichText", QXlsx::Format());
    sst.addSharedString(rs);

    for (int i=0; i<3; ++i) {
        QXlsx::RichString rs2;
        rs2.addFragment("Hello", QXlsx::Format());
        rs2.addFragment(" Qt!", QXlsx::Format());
        sst.addSharedString(rs2);
    }

    sst.addSharedString("Hello World");
    sst.addSharedString("Hello Qt!");

    QByteArray xmlData = sst.saveToXmlData();
    QXmlStreamReader reader(xmlData);

    int count = 0;
    int uniqueCount = 0;
    while(!reader.atEnd()) {
        QXmlStreamReader::TokenType token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("sst")) {
                QXmlStreamAttributes attributes = reader.attributes();
                count = attributes.value("count").toString().toInt();
                uniqueCount = attributes.value("uniqueCount").toString().toInt();
            }
        }
    }

    QCOMPARE(count, 8);
    QCOMPARE(uniqueCount, 5);
}

void SharedStringsTest::testRemoveSharedString()
{
    QXlsx::SharedStrings sst(QXlsx::SharedStrings::F_NewFromScratch);
    sst.addSharedString("Hello Qt!");
    sst.addSharedString("Xlsx Writer");
    sst.addSharedString("Hello World");
    sst.addSharedString("Hello Qt!");
    sst.addSharedString("Hello Qt!");

    sst.removeSharedString("Hello World");
    sst.removeSharedString("Hello Qt!");
    sst.removeSharedString("Non exists");

    QByteArray xmlData = sst.saveToXmlData();
    QXmlStreamReader reader(xmlData);

    int count = 0;
    int uniqueCount = 0;
    while(!reader.atEnd()) {
        QXmlStreamReader::TokenType token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("sst")) {
                QXmlStreamAttributes attributes = reader.attributes();
                count = attributes.value("count").toString().toInt();
                uniqueCount = attributes.value("uniqueCount").toString().toInt();
            }
        }
    }

    QCOMPARE(count, 3);
    QCOMPARE(uniqueCount, 2);
}

void SharedStringsTest::testLoadXmlData()
{
    QXlsx::SharedStrings sst(QXlsx::SharedStrings::F_NewFromScratch);
    sst.addSharedString("Hello Qt!");
    sst.addSharedString("Xlsx Writer");

    QXlsx::RichString rs;
    rs.addFragment("Hello", QXlsx::Format());
    rs.addFragment(" RichText", QXlsx::Format());
    sst.addSharedString(rs);
    for (int i=0; i<3; ++i) {
        QXlsx::RichString rs2;
        rs2.addFragment("Hello", QXlsx::Format());
        rs2.addFragment(" Qt!", QXlsx::Format());
        sst.addSharedString(rs2);
    }
    sst.addSharedString("Hello World");
    sst.addSharedString("Hello Qt!");
    QByteArray xmlData = sst.saveToXmlData();

    QSharedPointer<QXlsx::SharedStrings> sst2(new QXlsx::SharedStrings(QXlsx::SharedStrings::F_LoadFromExists));
    sst2->loadFromXmlData(xmlData);

    QCOMPARE(sst2->getSharedString(0).toPlainString(), QStringLiteral("Hello Qt!"));
    QCOMPARE(sst2->getSharedString(2), rs);
    QCOMPARE(sst2->getSharedString(4).toPlainString(), QStringLiteral("Hello World"));
    QCOMPARE(sst2->getSharedStringIndex("Hello Qt!"), 0);
    QCOMPARE(sst2->getSharedStringIndex("Hello World"), 4);
}

void SharedStringsTest::testLoadRichStringXmlData()
{
    QByteArray xmlData = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">"
            "<si>"
            "<r><t>e=mc</t></r>"
            "<r>"
            "<rPr><vertAlign val=\"superscript\"/>"
            "<sz val=\"11\"/>"
            "<rFont val=\"MyFontName\"/>"
            "<family val=\"3\"/>"
            "<charset val=\"134\"/>"
            "<scheme val=\"minor\"/>"
            "</rPr>"
            "<t>2</t></r>"
            "</si>"
            "</sst>";

    QSharedPointer<QXlsx::SharedStrings> sst(new QXlsx::SharedStrings(QXlsx::SharedStrings::F_LoadFromExists));
    sst->loadFromXmlData(xmlData);
    QXlsx::RichString rs = sst->getSharedString(0);
    QVERIFY(rs.fragmentText(0) == "e=mc");
    QVERIFY(rs.fragmentText(1) == "2");
    QVERIFY(rs.fragmentFormat(0) == QXlsx::Format());
    QXlsx::Format format = rs.fragmentFormat(1);
    QCOMPARE(format.fontName(), QString("MyFontName"));
    QCOMPARE(format.fontScript(), QXlsx::Format::FontScriptSuper);
    QCOMPARE(format.fontSize(), 11);
}

QTEST_APPLESS_MAIN(SharedStringsTest)

#include "tst_sharedstringstest.moc"
