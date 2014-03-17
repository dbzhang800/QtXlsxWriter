#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxworksheet.h"

QTXLSX_USE_NAMESPACE

int main()
{
    //![0]
    Document xlsx;
    //![0]

    //![1]
    xlsx.setColumnWidth(1, 2, 40);
    Format rAlign;
    rAlign.setHorizontalAlignment(Format::AlignRight);
    Format lAlign;
    lAlign.setHorizontalAlignment(Format::AlignLeft);
    xlsx.write("B3", 40, lAlign);
    xlsx.write("B4", 30, lAlign);
    xlsx.write("B5", 50, lAlign);
    xlsx.write("A7", "SUM(B3:B5)=", rAlign);
    xlsx.write("B7", "=SUM(B3:B5)", lAlign);
    xlsx.write("A8", "AVERAGE(B3:B5)=", rAlign);
    xlsx.write("B8", "=AVERAGE(B3:B5)", lAlign);
    xlsx.write("A9", "MAX(B3:B5)=", rAlign);
    xlsx.write("B9", "=MAX(B3:B5)", lAlign);
    xlsx.write("A10", "MIN(B3:B5)=", rAlign);
    xlsx.write("B10", "=MIN(B3:B5)", lAlign);
    xlsx.write("A11", "COUNT(B3:B5)=", rAlign);
    xlsx.write("B11", "=COUNT(B3:B5)", lAlign);

    xlsx.write("A13", "IF(B7>100,\"large\",\"small\")=", rAlign);
    xlsx.write("B13", "=IF(B7>100,\"large\",\"small\")", lAlign);

    xlsx.write("A15", "SQRT(25)=", rAlign);
    xlsx.write("B15", "=SQRT(25)", lAlign);
    xlsx.write("A16", "RAND()=", rAlign);
    xlsx.write("B16", "=RAND()", lAlign);
    xlsx.write("A17", "2*PI()=", rAlign);
    xlsx.write("B17", "=2*PI()", lAlign);

    xlsx.write("A19", "UPPER(\"qtxlsx\")=", rAlign);
    xlsx.write("B19", "=UPPER(\"qtxlsx\")", lAlign);
    xlsx.write("A20", "LEFT(\"ubuntu\",3)=", rAlign);
    xlsx.write("B20", "=LEFT(\"ubuntu\",3)", lAlign);
    xlsx.write("A21", "LEN(\"Hello Qt!\")=", rAlign);
    xlsx.write("B21", "=LEN(\"Hello Qt!\")", lAlign);
    //![1]

    //![2]
    xlsx.addSheet("ArrayFormula");
    Worksheet *sheet = xlsx.currentWorksheet();

    for (int row=2; row<20; ++row) {
        sheet->write(row, 2, QString(row%5+1, 'X')); //B2:B19
        sheet->write(row, 3, QString(row%5+1, 'X')); //C2:C19
        sheet->write(row, 5, 100.0 - row); //E2:E19
    }
    sheet->writeArrayFormula("C20", "{=SUM(IF((C2:C19=\"X\")*(B2:B19=\"X\"),1,0))}");
    sheet->writeArrayFormula("F2:F19", "{=E2:E19*10}");
    //![2]

    //![3]
    xlsx.save();
    //![3]

    return 0;
}
