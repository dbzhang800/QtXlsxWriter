#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxconditionalformatting.h"

using namespace QXlsx;

int main()
{
    //![0]
    Document xlsx;
    Format hFmt;
    hFmt.setFontBold(true);
    xlsx.write("B1", "(-inf,40)", hFmt);
    xlsx.write("C1", "[30,70]", hFmt);
    xlsx.write("D1", "startsWith 2", hFmt);
    xlsx.write("E1", "dataBar", hFmt);
    xlsx.write("F1", "colorScale", hFmt);

    for (int row=3; row<22; ++row) {
        for (int col=2; col<22; ++col)
            xlsx.write(row, col, qrand() % 100);
    }
    //![0]

    //![cf1]
    ConditionalFormatting cf1;
    Format fmt1;
    fmt1.setFontColor(Qt::green);
    fmt1.setBorderStyle(Format::BorderDashed);
    cf1.addHighlightCellsRule(ConditionalFormatting::Highlight_LessThan, "40", fmt1);
    cf1.addRange("B3:B21");
    xlsx.addConditionalFormatting(cf1);
    //![cf1]

    //![cf2]
    ConditionalFormatting cf2;
    Format fmt2;
    fmt2.setBorderStyle(Format::BorderDotted);
    fmt2.setBorderColor(Qt::blue);
    cf2.addHighlightCellsRule(ConditionalFormatting::Highlight_Between, "30", "70", fmt2);
    cf2.addRange("C3:C21");
    xlsx.addConditionalFormatting(cf2);
    //![cf2]

    //![cf3]
    ConditionalFormatting cf3;
    Format fmt3;
    fmt3.setFontStrikeOut(true);
    fmt3.setFontBold(true);
    cf3.addHighlightCellsRule(ConditionalFormatting::Highlight_BeginsWith, "2", fmt3);
    cf3.addRange("D3:D21");
    xlsx.addConditionalFormatting(cf3);
    //![cf3]

    //![cf4]
    ConditionalFormatting cf4;
    cf4.addDataBarRule(Qt::blue);
    cf4.addRange("E3:E21");
    xlsx.addConditionalFormatting(cf4);
    //![cf4]

    //![cf5]
    ConditionalFormatting cf5;
    cf5.add2ColorScaleRule(Qt::blue, Qt::red);
    cf5.addRange("F3:F21");
    xlsx.addConditionalFormatting(cf5);
    //![cf5]

    //![2]
    xlsx.saveAs("Book1.xlsx");
    //![2]

    Document xlsx2("Book1.xlsx");
    xlsx2.saveAs("Book2.xlsx");
    return 0;
}
