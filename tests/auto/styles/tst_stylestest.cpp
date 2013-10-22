#include "private/xlsxstyles_p.h"
#include "private/xlsxxmlreader_p.h"
#include "xlsxformat.h"
#include "private/xlsxformat_p.h"
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
    void testAddFormat2();
    void testSolidFillBackgroundColor();

    void testReadFonts();
    void testReadFills();
    void testReaderBorders();
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

void StylesTest::testAddFormat2()
{
    QXlsx::Styles styles;

    QXlsx::Format *format = styles.createFormat();
    format->setNumberFormat("h:mm:ss AM/PM"); //builtin 19
    styles.addFormat(format);

    QCOMPARE(format->numberFormatIndex(), 19);

    QXlsx::Format *format2 = styles.createFormat();
    format2->setNumberFormat("aaaaa h:mm:ss AM/PM"); //custom
    styles.addFormat(format2);

    QCOMPARE(format2->numberFormatIndex(), 164);
}

// For a solid fill, Excel reverses the role of foreground and background colours
void StylesTest::testSolidFillBackgroundColor()
{
    QXlsx::Styles styles;
    QXlsx::Format *format = styles.createFormat();
    format->setPatternBackgroundColor(QColor(Qt::red));
    styles.addFormat(format);

    QByteArray xmlData = styles.saveToXmlData();

    QVERIFY(xmlData.contains("<patternFill patternType=\"solid\"><fgColor rgb=\"FFff0000\"/>"));
}

void StylesTest::testReadFonts()
{
    QByteArray xmlData = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<fonts count=\"3\">"
            "<font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"
            "<font><sz val=\"15\"/><color rgb=\"FFff0000\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"
            "<font><b/><u val=\"double\"/><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"
            "</fonts>";
    QXlsx::Styles styles(true);
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//So current node is fonts
    styles.readFonts(reader);

    QCOMPARE(styles.m_fontsList.size(), 3);
}

void StylesTest::testReadFills()
{
    QByteArray xmlData = "<fills count=\"4\">"
            "<fill><patternFill patternType=\"none\"/></fill>"
            "<fill><patternFill patternType=\"gray125\"/></fill>"
            "<fill><patternFill patternType=\"lightUp\"/></fill>"
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFa0a0a4\"/></patternFill></fill>"
            "</fills>";
    QXlsx::Styles styles(true);
    QXlsx::XmlStreamReader reader(xmlData);
    reader.readNextStartElement();//So current node is fonts
    styles.readFills(reader);

    QCOMPARE(styles.m_fillsList.size(), 4);
    QCOMPARE(styles.m_fillsList[3]->pattern, QXlsx::Format::PatternSolid);
    QCOMPARE(styles.m_fillsList[3]->bgColor, QColor(Qt::gray));//for solid pattern, bg vs. fg color!
}

void StylesTest::testReaderBorders()
{

}

QTEST_APPLESS_MAIN(StylesTest)

#include "tst_stylestest.moc"
