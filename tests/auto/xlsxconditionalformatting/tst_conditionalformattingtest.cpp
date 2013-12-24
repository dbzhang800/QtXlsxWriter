#include "xlsxconditionalformatting.h"
#include "xlsxformat.h"
#include "private/xlsxconditionalformatting_p.h"

#include <QString>
#include <QtTest>
#include <QBuffer>
#include <QXmlStreamWriter>

using namespace QXlsx;

class ConditionalFormattingTest : public QObject
{
    Q_OBJECT

public:
    ConditionalFormattingTest();

private Q_SLOTS:
    void testHighlightRules();
    void testHighlightRules_data();
    void testDataBarRules();
};

ConditionalFormattingTest::ConditionalFormattingTest()
{
}

void ConditionalFormattingTest::testHighlightRules_data()
{
    QTest::addColumn<int>("type");
    QTest::addColumn<QString>("formula1");
    QTest::addColumn<QString>("formula2");
    QTest::addColumn<QByteArray>("result");

    QTest::newRow("lessThan")<<(int)ConditionalFormatting::Highlight_LessThan
                            <<"100"
                            <<QString()
                            <<QByteArray("<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"lessThan\"><formula>100</formula></cfRule>");
    QTest::newRow("between")<<(int)ConditionalFormatting::Highlight_Between
                              <<"4"
                              <<"20"
                              <<QByteArray("<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"between\"><formula>4</formula><formula>20</formula></cfRule>");

    QTest::newRow("containsText")<<(int)ConditionalFormatting::Highlight_ContainsText
                                 <<"Qt"
                                 <<QString()
                                 <<QByteArray("<cfRule type=\"containsText\" dxfId=\"0\" priority=\"1\" operator=\"containsText\" text=\"Qt\">");
    QTest::newRow("beginsWith")<<(int)ConditionalFormatting::Highlight_BeginsWith
                                 <<"Qt"
                                 <<QString()
                                 <<QByteArray("<cfRule type=\"beginsWith\" dxfId=\"0\" priority=\"1\" operator=\"beginsWith\" text=\"Qt\"><formula>LEFT(C3,LEN"); //(\"Qt\"))=\"Qt\"</formula></cfRule>");
    QTest::newRow("duplicateValues")<<(int)ConditionalFormatting::Highlight_Duplicate
                            <<QString()
                            <<QString()
                            <<QByteArray("<cfRule type=\"duplicateValues\" dxfId=\"0\" priority=\"1\"/>");
}

void ConditionalFormattingTest::testHighlightRules()
{
    QFETCH(int, type);
    QFETCH(QString, formula1);
    QFETCH(QString, formula2);
    QFETCH(QByteArray, result);

    Format fmt;
    fmt.setFontBold(true);
    fmt.setDxfIndex(0);

    ConditionalFormatting cf;
    cf.addHighlightCellsRule((ConditionalFormatting::HighlightRuleType)type, formula1, formula2, fmt);
    cf.addRange("C3:C10");

    QBuffer buffer;
    buffer.open(QIODevice::WriteOnly);
    QXmlStreamWriter writer(&buffer);
    cf.saveToXml(writer);
    qDebug()<<buffer.buffer();
    QVERIFY(buffer.buffer().contains(result));
}

void ConditionalFormattingTest::testDataBarRules()
{
    ConditionalFormatting cf;
    cf.addDataBarRule(Qt::blue);
    cf.addRange("C3:C10");

    QBuffer buffer;
    buffer.open(QIODevice::WriteOnly);
    QXmlStreamWriter writer(&buffer);
    cf.saveToXml(writer);
    qDebug()<<buffer.buffer();
    QByteArray res = "<cfRule type=\"dataBar\" priority=\"1\">"
            "<dataBar><cfvo type=\"min\" val=\"0\"/>"
            "<cfvo type=\"max\" val=\"0\"/>"
            "<color rgb=\"FF0000FF\"/></dataBar>"
            "</cfRule>";
    QVERIFY(buffer.buffer().contains(res));
}

QTEST_APPLESS_MAIN(ConditionalFormattingTest)

#include "tst_conditionalformattingtest.moc"
