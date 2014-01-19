#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

using namespace QXlsx;

int main()
{
    //![0]
    Document xlsx;
    for (int i=1; i<10; ++i)
        xlsx.write(i, 1, i*i);
    //![0]

    //![1]
    Chart *pieChart = xlsx.insertChart(3, 3, QSize(300, 300));
    pieChart->setChartType(Chart::CT_Pie);
    pieChart->addSeries(CellRange("A1:A9"));

    Chart *pie3DChart = xlsx.insertChart(3, 7, QSize(300, 300));
    pie3DChart->setChartType(Chart::CT_Pie3D);
    pie3DChart->addSeries(CellRange("A1:A9"));

    Chart *barChart = xlsx.insertChart(23, 3, QSize(300, 300));
    barChart->setChartType(Chart::CT_Bar);
    barChart->addSeries(CellRange("A1:A9"));

    Chart *bar3DChart = xlsx.insertChart(23, 7, QSize(300, 300));
    bar3DChart->setChartType(Chart::CT_Bar3D);
    bar3DChart->addSeries(CellRange("A1:A9"));

    //![1]

    //![2]
    xlsx.saveAs("Book1.xlsx");
    //![2]

    Document xlsx2("Book1.xlsx");
    xlsx2.saveAs("Book2.xlsx");
    return 0;
}
