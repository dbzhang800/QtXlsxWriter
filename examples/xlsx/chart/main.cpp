#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxchart.h"
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
    Chart *pieChart = sheet->insertChart(3, 3, QSize(300, 300));
    pieChart->setChartType(Chart::CT_Pie);
    pieChart->addSeries(CellRange("A1:A9"));

    Chart *barChart = sheet->insertChart(6, 6, QSize(300, 300));
    barChart->setChartType(Chart::CT_Bar);
    barChart->addSeries(CellRange("A1:A9"));

    //![1]

    //![2]
    xlsx.saveAs("Book1.xlsx");
    //![2]

    Document xlsx2("Book1.xlsx");
    xlsx2.saveAs("Book2.xlsx");
    return 0;
}
