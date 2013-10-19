#include "xlsxdocument.h"

int main()
{
    //Generate a simple xlsx file at first.
    QXlsx::Document xlsx;
    xlsx.write("A1", "Hello Qt!");
    xlsx.write("A2", 500);
    xlsx.saveAs("first.xlsx");

    //Read, edit, save
    QXlsx::Document xlsx2("first.xlsx");
    xlsx2.addWorksheet("Second");
    xlsx2.write("A1", "Hello Qt again!");
    xlsx2.saveAs("second.xlsx");

    return 0;
}
