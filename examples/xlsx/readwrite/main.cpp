#include "xlsxdocument.h"
#include "xlsxformat.h"

int main()
{
    //Generate a simple xlsx file at first.
    //![0]
    QXlsx::Document xlsx;
    xlsx.setDocumentProperty("title", "This is an example spreadsheet");
    xlsx.setDocumentProperty("creator", "Qt Xlsx Library");
    xlsx.setSheetName("First Sheet");
    QXlsx::Format *format = xlsx.createFormat();
    format->setFontColor(QColor(Qt::blue));
    format->setFontSize(15);
    xlsx.write("A1", "Hello Qt!", format);
    xlsx.write("A2", 500);
    xlsx.saveAs("first.xlsx");
    //![0]

    //Read, edit, save
    //![1]
    QXlsx::Document xlsx2("first.xlsx");
    xlsx2.write("A3", "Hello Qt again!");
    xlsx2.addWorksheet("Second Sheet");
    xlsx2.write("A1", "Hello Qt again!");
    xlsx2.setCurrentWorksheet(0);
    xlsx2.saveAs("second.xlsx");
    //![1]

    return 0;
}
