#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

using namespace QXlsx;

int main()
{
    //![0]
    Document xlsx;

    int binSize = 5;
    int binCount = 11;
    int gaussianCount = 21;

    // histogram
    for (int i=0; i<binCount; ++i) {
        xlsx.write(i + 1, 1, QString("%1 - %2").arg(i * binSize).arg((i + 1) * binSize - 1)); // A1:A11
    }
    for (int i=0; i<binCount; ++i) {
        xlsx.write(i + 1, 2, - (i - binCount / 2) * (i - binCount / 2) + (binCount * binCount) / 4); // B1:B11
    }

    // gaussian
    float step = ((float)binCount * (float)binSize - 1.0) / (gaussianCount - 1);
    for (int i=0; i<gaussianCount; ++i) {
        xlsx.write(i + 1, 3, (float)i * step); // C1:C21
        xlsx.write(i + 1, 4, - ((float)i - (float)gaussianCount / 2) * ((float)i - (float)gaussianCount / 2) / (((float)gaussianCount * (float)gaussianCount) / 4) + 1); // D1:D21
    }
    //![0]

    //![1]
    Chart *histogramChart = xlsx.insertChart(3, 9, QSize(300, 300));
    histogramChart->setChartType(Chart::CT_Histogram);
    histogramChart->addSeries(CellRange("A1:B11"));

    Chart *histogramGaussianChart = xlsx.insertChart(23, 9, QSize(300, 300));
    histogramGaussianChart->setChartType(Chart::CT_Histogram);
    histogramGaussianChart->addSeries(CellRange("A1:B11"));
    histogramGaussianChart->addSeries(CellRange("C1:D21"));
    //![1]

    //![2]
    xlsx.saveAs("Book1.xlsx");
    //![2]

    return 0;
}
