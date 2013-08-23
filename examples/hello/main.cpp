#include <QtCore>
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"

int main(int argc, char* argv[])
{
#ifdef Q_OS_MAC
    QXlsx::Workbook workbook("../../../Test.xlsx");
#else
    QXlsx::Workbook workbook("Test.xlsx");
#endif
    QXlsx::Worksheet *sheet = workbook.addWorksheet();
    sheet->write("A1", "Hello Qt!");
    workbook.close();
    return 0;
}
