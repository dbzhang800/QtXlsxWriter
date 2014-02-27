#include "private/xlsxdocpropsapp_p.h"
#include <QString>
#include <QtTest>

class DocPropsAppTest : public QObject
{
    Q_OBJECT
    
public:
    DocPropsAppTest();
    
private Q_SLOTS:
    void testCase1();
};

DocPropsAppTest::DocPropsAppTest()
{

}

void DocPropsAppTest::testCase1()
{
    QXlsx::DocPropsApp props(QXlsx::DocPropsApp::F_NewFromScratch);

    props.setProperty("company", "HMI CN");
    props.setProperty("manager", "Debao");

    QFile f1("temp.xml");
    f1.open(QFile::WriteOnly);
    props.saveToXmlFile(&f1);
    f1.close();

    f1.open(QFile::ReadOnly);
    QXlsx::DocPropsApp props2(QXlsx::DocPropsApp::F_LoadFromExists);
    props2.loadFromXmlFile(&f1);

    QCOMPARE(props2.property("company"), QString("HMI CN"));
    QCOMPARE(props2.property("manager"), QString("Debao"));
    QFile::remove("temp.xml");
}

QTEST_APPLESS_MAIN(DocPropsAppTest)

#include "tst_docpropsapptest.moc"
