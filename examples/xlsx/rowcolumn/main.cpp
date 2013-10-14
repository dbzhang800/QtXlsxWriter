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
    xlsx.write(0, 2, "Row:0, Col:2 ==> (C1)");

    //Set the height of the first row to 50.0(points)
    xlsx.setRow(0, 50.0);

    //Set the width of the third column to 40.0(chars)
    xlsx.setColumn(2, 3, 40.0);

    //Set style for the row 11th.
    QXlsx::Format *format1 = xlsx.createFormat();
    format1->setFontBold(true);
    format1->setFontColor(QColor(Qt::blue));
    format1->setFontSize(20);
    xlsx.write(10, 0, "Hello Row Style");
    xlsx.write(10, 5, "Blue Color");
    xlsx.setRow(10, 40, format1);

    //Set style for the col [9th, 16th)
    QXlsx::Format *format2 = xlsx.createFormat();
    format2->setFontBold(true);
    format2->setFontColor(QColor(Qt::magenta));
    for (int row=11; row<30; row++)
        for (int col=8; col<15; col++)
            xlsx.write(row, col, row+col);
    xlsx.setColumn(8, 15, 5.0, format2);

    xlsx.saveAs(DATA_PATH"Test.xlsx");
    return 0;
}
