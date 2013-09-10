#include <QtCore>
#include "xlsxdocument.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main()
{
    QXlsx::Document xlsx;

    //Write to first worksheet.
    xlsx.write("A1", "Hello Qt!");
    xlsx.write("B3", 12345);
    xlsx.write("C5", "=44+33");
    xlsx.write("D7", true);
    xlsx.write("E1", "http://qt-project.org");

    //Create another worksheet.
    xlsx.addWorksheet();
    //Rows and columns are zero indexed.
    //The first cell in a worksheet, "A1", is (0, 0).
    xlsx.write(0, 0, "First");
    xlsx.write(1, 0, "Second");
    xlsx.write(2, 0, "Third");
    xlsx.write(3, 0, "Fourth");
    xlsx.write(4, 0, "Total");
    xlsx.write(0, 1, 100);
    xlsx.write(1, 1, 200);
    xlsx.write(2, 1, 300);
    xlsx.write(3, 1, 400);
    xlsx.write(4, 1, "=SUM(B1:B4)");

    xlsx.saveAs(DATA_PATH"Test.xlsx");
    return 0;
}
