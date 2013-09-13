#include "private/xlsxstyles_p.h"
#include "xlsxformat.h"
#include <QString>
#include <QtTest>

class StylesTest : public QObject
{
    Q_OBJECT

public:
    StylesTest();

private Q_SLOTS:
    void testEmptyStyle();
    void testAddFormat();
};

StylesTest::StylesTest()
{
}

void StylesTest::testEmptyStyle()
{
    QXlsx::Styles styles;
    QByteArray xmlData = styles.saveToXmlData();

    QVERIFY2(xmlData.contains("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"), "Must have one cell style");
}

void StylesTest::testAddFormat()
{
    QXlsx::Styles styles;

    for (int i=0; i<10; ++i) {
        QXlsx::Format *format = styles.createFormat();
        format->setFontBold(true);
        styles.addFormat(format);
    }

    QByteArray xmlData = styles.saveToXmlData();
    QVERIFY2(xmlData.contains("<cellXfs count=\"2\">"), ""); //Note we have a default one
}

QTEST_APPLESS_MAIN(StylesTest)

#include "tst_stylestest.moc"
