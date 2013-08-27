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

    workbook.save(DATA_PATH"TestStyle.xlsx");
    workbook.save(DATA_PATH"TestStyle.zip");
    return 0;
}
