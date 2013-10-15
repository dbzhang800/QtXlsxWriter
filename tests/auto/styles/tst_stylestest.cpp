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
    void testAddFormat2();
    void testSolidFillBackgroundColor();
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

QTEST_APPLESS_MAIN(StylesTest)

#include "tst_stylestest.moc"
