#include <QtCore>
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main()
{
    QXlsx::Workbook workbook;
    QXlsx::Worksheet *sheet = workbook.addWorksheet();
    sheet->write("A1", "Hello Qt!");
    sheet->write("B3", 12345);
    sheet->write("C5", "=44+33");
    sheet->write("D7", true);

    QXlsx::Worksheet *sheet2 = workbook.addWorksheet();
    //Rows and columns are zero indexed.
    //The first cell in a worksheet, "A1", is (0, 0).
    sheet2->write(0, 0, "First");
    sheet2->write(1, 0, "Second");
    sheet2->write(2, 0, "Third");
    sheet2->write(3, 0, "Fourth");
    sheet2->write(4, 0, "Total");
    sheet2->write(0, 1, 100);
    sheet2->write(1, 1, 200);
    sheet2->write(2, 1, 300);
    sheet2->write(3, 1, 400);
    sheet2->write(4, 1, "=SUM(B1:B4)");

    workbook.setActivedWorksheet(1);

    workbook.save(DATA_PATH"Test.xlsx");
    workbook.save(DATA_PATH"Test.zip");
    return 0;
}
