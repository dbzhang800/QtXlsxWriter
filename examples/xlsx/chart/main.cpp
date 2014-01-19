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

    Chart *lineChart = xlsx.insertChart(43, 3, QSize(300, 300));
    lineChart->setChartType(Chart::CT_Line);
    lineChart->addSeries(CellRange("A1:A9"));

    Chart *line3DChart = xlsx.insertChart(43, 7, QSize(300, 300));
    line3DChart->setChartType(Chart::CT_Line3D);
    line3DChart->addSeries(CellRange("A1:A9"));

    Chart *areaChart = xlsx.insertChart(63, 3, QSize(300, 300));
    areaChart->setChartType(Chart::CT_Area);
    areaChart->addSeries(CellRange("A1:A9"));

    Chart *area3DChart = xlsx.insertChart(63, 7, QSize(300, 300));
    area3DChart->setChartType(Chart::CT_Area3D);
    area3DChart->addSeries(CellRange("A1:A9"));

    Chart *scatterChart = xlsx.insertChart(83, 3, QSize(300, 300));
    scatterChart->setChartType(Chart::CT_Scatter);
    scatterChart->addSeries(CellRange("A1:A9"));

    Chart *doughnutChart = xlsx.insertChart(103, 3, QSize(300, 300));
    doughnutChart->setChartType(Chart::CT_Doughnut);
    doughnutChart->addSeries(CellRange("A1:A9"));
    //![1]

    //![2]
    xlsx.saveAs("Book1.xlsx");
    //![2]

    Document xlsx2("Book1.xlsx");
    xlsx2.saveAs("Book2.xlsx");
    return 0;
}
