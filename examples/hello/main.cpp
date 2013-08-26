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
    workbook.save(DATA_PATH"Test.xlsx");
    workbook.save(DATA_PATH"Test.zip");
    return 0;
}
