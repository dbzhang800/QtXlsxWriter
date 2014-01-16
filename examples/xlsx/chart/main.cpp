#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxpiechart.h"
#include "xlsxworksheet.h"

using namespace QXlsx;

int main()
{
    //![0]
    Document xlsx;

    Worksheet *sheet = xlsx.currentWorksheet();
    for (int i=1; i<10; ++i)
        sheet->write(i, 1, i*i);
    //![0]

    //![1]
    PieChart *chart = new PieChart;
    chart->addSeries(CellRange("A1:A9"), sheet->sheetName());
    sheet->insertChart(3, 3, chart, QSize(300, 300));
    //![1]

    //![2]
    xlsx.saveAs("Book1.xlsx");
    //![2]

    Document xlsx2("Book1.xlsx");
    xlsx2.saveAs("Book2.xlsx");
    return 0;
}
