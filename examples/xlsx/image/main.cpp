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
    QImage image(400, 300, QImage::Format_RGB32);
    image.fill(Qt::green);
    sheet->insertImage(5, 5, image);

    workbook.save(DATA_PATH"Test.xlsx");
//    workbook.save(DATA_PATH"Test2.zip");

    return 0;
}
