#include <QtCore>
#include "xlsxdocument.h"

int main()
{
    //![0]
    QXlsx::Document xlsx;
    //![0]

    //![1]
    xlsx.write("A1", "Hello Qt!");
    xlsx.write("A2", 12345);
    xlsx.write("A3", "=44+33");
    xlsx.write("A4", true);
    xlsx.write("A5", "http://qt-project.org");
    //![1]

    //![2]
    xlsx.save();
    //![2]

    QXlsx::Document xlsx2;
    xlsx2.addWorksheet();
    xlsx2.saveAs("Book2.xlsx");

    return 0;
}
