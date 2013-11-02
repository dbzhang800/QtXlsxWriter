#include "private/xlsxsharedstrings_p.h"
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
                count = attributes.value("count").toInt();
                uniqueCount = attributes.value("uniqueCount").toInt();
            }
        }
    }

    QCOMPARE(count, 4);
    QCOMPARE(uniqueCount, 3);
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
                count = attributes.value("count").toInt();
                uniqueCount = attributes.value("uniqueCount").toInt();
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
    sst.addSharedString("Hello World");
    sst.addSharedString("Hello Qt!");
    QByteArray xmlData = sst.saveToXmlData();

    QSharedPointer<QXlsx::SharedStrings> sst2 = QXlsx::SharedStrings::loadFromXmlData(xmlData);

    QCOMPARE(sst2->getSharedString(0), QStringLiteral("Hello Qt!"));
    QCOMPARE(sst2->getSharedString(2), QStringLiteral("Hello World"));
    QCOMPARE(sst2->getSharedStringIndex("Hello Qt!"), 0);
    QCOMPARE(sst2->getSharedStringIndex("Hello World"), 2);
}

QTEST_APPLESS_MAIN(SharedStringsTest)

#include "tst_sharedstringstest.moc"
