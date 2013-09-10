#include <QtGui>
#include "xlsxdocument.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main(int argc, char** argv)
{
    QGuiApplication(argc, argv);

    QXlsx::Document xlsx;

    xlsx.write("B1", "Merge Cells");
    xlsx.mergeCells("B1:B5");

    xlsx.write("E2", "Merge Cells 2");
    xlsx.mergeCells("E2:G4");

    xlsx.saveAs(DATA_PATH"Test.xlsx");

    return 0;
}

