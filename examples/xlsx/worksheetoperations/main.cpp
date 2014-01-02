#include <QtCore>
#include "xlsxdocument.h"

int main()
{
    QXlsx::Document xlsx;

    xlsx.renameWorksheet("Sheet1", "TheFirstSheet");

    for (int i=1; i<20; ++i) {
        for (int j=1; j<15; ++j)
            xlsx.write(i, j, QString("R %1 C %2").arg(i).arg(j));
    }

    xlsx.addWorksheet("TheSecondSheet");
    xlsx.write(2, 2, "Hello Qt Xlsx");

    xlsx.copyWorksheet("TheFirstSheet", "CopyOfTheFirst");

    xlsx.addWorksheet("TheForthSheet");
    xlsx.write(3, 3, "This will be deleted...");

    xlsx.selectWorksheet("CopyOfTheFirst");
    xlsx.write(25, 2, "On the Copy Sheet");

    xlsx.deleteWorksheet("TheForthSheet");

    xlsx.moveWorksheet("TheSecondSheet", 0);

    xlsx.save();

    return 0;
}
