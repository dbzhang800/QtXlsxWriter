#include "private/xlsxsharedstrings_p.h"
#include "private/xlsxrichstring_p.h"
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

};

SharedStringsTest::SharedStringsTest()
{
}

void SharedStringsTest::testAddSharedString()
{
    QXlsx::SharedStrings sst;
    sst.addSharedString("Hello Qt!");
    sst.addSharedString("Xlsx Writer");

    QXlsx::RichString rs;
    rs.addFragment("Hello", 0);
    rs.addFragment(" RichText", 0);
    sst.addSharedString(rs);

    for (int i=0; i<3; ++i) {
        QXlsx::RichString rs2;
        rs2.addFragment("Hello", 0);
        rs2.addFragment(" Qt!", 0);
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
    QXlsx::SharedStrings sst;
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
    QXlsx::SharedStrings sst;
    sst.addSharedString("Hello Qt!");
    sst.addSharedString("Xlsx Writer");

    QXlsx::RichString rs;
    rs.addFragment("Hello", 0);
    rs.addFragment(" RichText", 0);
    sst.addSharedString(rs);
    for (int i=0; i<3; ++i) {
        QXlsx::RichString rs2;
        rs2.addFragment("Hello", 0);
        rs2.addFragment(" Qt!", 0);
        sst.addSharedString(rs2);
    }
    sst.addSharedString("Hello World");
    sst.addSharedString("Hello Qt!");
    QByteArray xmlData = sst.saveToXmlData();

    QSharedPointer<QXlsx::SharedStrings> sst2(new QXlsx::SharedStrings);
    sst2->loadFromXmlData(xmlData);

    QCOMPARE(sst2->getSharedString(0).toPlainString(), QStringLiteral("Hello Qt!"));
    QCOMPARE(sst2->getSharedString(2), rs);
    QCOMPARE(sst2->getSharedString(4).toPlainString(), QStringLiteral("Hello World"));
    QCOMPARE(sst2->getSharedStringIndex("Hello Qt!"), 0);
    QCOMPARE(sst2->getSharedStringIndex("Hello World"), 4);
}

QTEST_APPLESS_MAIN(SharedStringsTest)

#include "tst_sharedstringstest.moc"
