#include "private/xlsxdocpropscore_p.h"
#include <QString>
#include <QtTest>
#include <QFile>

class DocPropsCoreTest : public QObject
{
    Q_OBJECT
    
public:
    DocPropsCoreTest();
    
private Q_SLOTS:
    void testCase1();
};

DocPropsCoreTest::DocPropsCoreTest()
{
}

void DocPropsCoreTest::testCase1()
{
    QXlsx::DocPropsCore props(QXlsx::DocPropsCore::F_NewFromScratch);

    props.setProperty("creator", "Debao");
    props.setProperty("keywords", "Test, test, TEST");
    props.setProperty("title", "ABC");

    QFile f1("temp.xml");
    f1.open(QFile::WriteOnly);
    props.saveToXmlFile(&f1);
    f1.close();

    f1.open(QFile::ReadOnly);
    QXlsx::DocPropsCore props2(QXlsx::DocPropsCore::F_LoadFromExists);
    props2.loadFromXmlFile(&f1);

    QCOMPARE(props2.property("creator"), QString("Debao"));
    QCOMPARE(props2.property("keywords"), QString("Test, test, TEST"));
    QCOMPARE(props2.property("title"), QString("ABC"));
    QFile::remove("temp.xml");
}

QTEST_APPLESS_MAIN(DocPropsCoreTest)

#include "tst_propscoretest.moc"
