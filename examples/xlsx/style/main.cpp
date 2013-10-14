#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxformat.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main()
{
    QXlsx::Document xlsx;
    QXlsx::Format *format1 = xlsx.createFormat();
    format1->setFontColor(QColor(Qt::red));
    format1->setFontSize(15);
    format1->setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    format1->setBorderStyle(QXlsx::Format::BorderDashDotDot);
    xlsx.write("A1", "Hello Qt!", format1);
    xlsx.write("B3", 12345, format1);

    QXlsx::Format *format2 = xlsx.createFormat();
    format2->setFontBold(true);
    format2->setFontUnderline(QXlsx::Format::FontUnderlineDouble);
    format2->setFillPattern(QXlsx::Format::PatternLightUp);
    xlsx.write("C5", "=44+33", format2);
    xlsx.write("D7", true, format2);

    QXlsx::Format *format3 = xlsx.createFormat();
    format3->setFontBold(true);
    format3->setFontColor(QColor(Qt::blue));
    format3->setFontSize(20);
    xlsx.write(10, 0, "Hello Row Style");
    xlsx.write(10, 5, "Blue Color");
    xlsx.setRow(10, 40, format3);

    QXlsx::Format *format4 = xlsx.createFormat();
    format4->setFontBold(true);
    format4->setFontColor(QColor(Qt::magenta));
    for (int row=20; row<40; row++)
        for (int col=8; col<15; col++)
            xlsx.write(row, col, row+col);
    xlsx.setColumn(8, 15, 5.0, format4);

    QXlsx::Format *format5 = xlsx.createFormat();
    format5->setNumberFormat(22);
    xlsx.write("A5", QDate(2013, 8, 29), format5);

    QXlsx::Format *format6 = xlsx.createFormat();
    format6->setPatternBackgroundColor(QColor(Qt::gray));
    xlsx.write("A6", "Background color: green", format6);

    xlsx.saveAs(DATA_PATH"TestStyle.xlsx");
    return 0;
}
