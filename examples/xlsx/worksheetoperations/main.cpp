#include <QtCore>
#include "xlsxdocument.h"

int main()
{
    QXlsx::Document xlsx;

    xlsx.renameSheet("Sheet1", "TheFirstSheet");

    for (int i=1; i<20; ++i) {
        for (int j=1; j<15; ++j)
            xlsx.write(i, j, QString("R %1 C %2").arg(i).arg(j));
    }

    xlsx.addSheet("TheSecondSheet");
    xlsx.write(2, 2, "Hello Qt Xlsx");

    xlsx.copySheet("TheFirstSheet", "CopyOfTheFirst");

    xlsx.addSheet("TheForthSheet");
    xlsx.write(3, 3, "This will be deleted...");

    xlsx.selectSheet("CopyOfTheFirst");
    xlsx.write(25, 2, "On the Copy Sheet");

    xlsx.deleteSheet("TheForthSheet");

    xlsx.moveSheet("TheSecondSheet", 0);

    xlsx.save();

    return 0;
}
