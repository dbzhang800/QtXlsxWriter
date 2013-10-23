#include "xlsxdocument.h"

int main()
{
    QXlsx::Document xlsx;

    xlsx.write("B1", "Merge Cells");
    xlsx.mergeCells("B1:B5");

    xlsx.write("E2", "Merge Cells 2");
    xlsx.mergeCells("E2:G4");

    xlsx.save();

    return 0;
}

