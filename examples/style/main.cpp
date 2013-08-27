#include <QtCore>
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"
#include "xlsxformat.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main()
{
    QXlsx::Workbook workbook;
    QXlsx::Worksheet *sheet = workbook.addWorksheet();

    QXlsx::Format *format1 = workbook.addFormat();
    format1->setFontColor(QColor(Qt::red));
    format1->setFontSize(15);
    sheet->write("A1", "Hello Qt!", format1);
    sheet->write("B3", 12345, format1);

    QXlsx::Format *format2 = workbook.addFormat();
    format2->setFontBold(true);
    format2->setFontUnderline(QXlsx::Format::FontUnderlineDouble);
    sheet->write("C5", "=44+33", format2);
    sheet->write("D7", true, format2);

    QXlsx::Format *format3 = workbook.addFormat();
    format3->setFontBold(true);
    format3->setFontColor(QColor(Qt::blue));
    format3->setFontSize(20);
    sheet->write(10, 0, "Hello Row Style");
    sheet->write(10, 5, "Blue Color");
    sheet->setRow(10, 40, format3);

    QXlsx::Format *format4 = workbook.addFormat();
    format4->setFontBold(true);
    format4->setFontColor(QColor(Qt::magenta));
    for (int row=20; row<40; row++)
        for (int col=8; col<15; col++)
            sheet->write(row, col, row+col);
    sheet->setColumn(8, 15, 5.0, format4);

    workbook.save(DATA_PATH"TestStyle.xlsx");
    return 0;
}
