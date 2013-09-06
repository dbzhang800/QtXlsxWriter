#include <QtGui>
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main(int argc, char** argv)
{
    QGuiApplication(argc, argv);

    QXlsx::Workbook workbook;
    QXlsx::Worksheet *sheet = workbook.addWorksheet();

    sheet->write("B1", "Merge Cells");
    sheet->mergeCells("B1:B5");

    sheet->write("E2", "Merge Cells 2");
    sheet->mergeCells("E2:G4");

    workbook.save(DATA_PATH"Test.xlsx");

    return 0;
}

