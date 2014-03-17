#include <QtCore>
#include "xlsxdocument.h"

int main()
{
    QXlsx::Document xlsx;
    QString sheetName="200";
    QString cellString;
    xlsx.addSheet(sheetName);
    for (int i=0;i<2000;i++) {
        for (int j=0;j<2000;++j)
        {
            cellString.sprintf("%d:%d",i,j);
            xlsx.write(i,j,cellString);
        }
    }

    xlsx.saveAs("Book1.xlsx");

    return 0;
}
