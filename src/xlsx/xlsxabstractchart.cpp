/****************************************************************************
** Copyright (c) 2013-2014 Debao Zhang <hello@debao.me>
** All right reserved.
**
** Permission is hereby granted, free of charge, to any person obtaining
** a copy of this software and associated documentation files (the
** "Software"), to deal in the Software without restriction, including
** without limitation the rights to use, copy, modify, merge, publish,
** distribute, sublicense, and/or sell copies of the Software, and to
** permit persons to whom the Software is furnished to do so, subject to
** the following conditions:
**
** The above copyright notice and this permission notice shall be
** included in all copies or substantial portions of the Software.
**
** THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
** EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
** MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
** NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
** LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
** OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
** WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
**
****************************************************************************/

#include "xlsxabstractchart.h"
#include "xlsxabstractchart_p.h"
#include "xlsxchartfile_p.h"
#include "xlsxcellrange.h"
#include "xlsxutility_p.h"

QT_BEGIN_NAMESPACE_XLSX

AbstractChartPrivate::AbstractChartPrivate(AbstractChart *chart)
    :q_ptr(chart)
{

}

AbstractChartPrivate::~AbstractChartPrivate()
{

}

/*!
 * \class AbstractChart
 *
 * Base class for all the charts.
 */


AbstractChart::AbstractChart()
    :d_ptr(new AbstractChartPrivate(this))
{
}

AbstractChart::AbstractChart(AbstractChartPrivate *d)
    :d_ptr(d)
{

}

AbstractChart::~AbstractChart()
{
    Q_D(AbstractChart);
    if (d->cf)
        d->cf->m_chart = 0;
}

void AbstractChart::addSeries(const CellRange &range, const QString &sheet)
{
    Q_D(AbstractChart);

    QString serRef = sheet;
    serRef += QLatin1String("!");
    serRef += xl_rowcol_to_cell(range.firstRow(), range.firstColumn(), true, true);
    serRef += QLatin1String(":");
    serRef += xl_rowcol_to_cell(range.lastRow(), range.lastColumn(), true, true);

    XlsxSeries *series = new XlsxSeries;
    series->numRef = serRef;

    d->seriesList.append(QSharedPointer<XlsxSeries>(series));
}

QT_END_NAMESPACE_XLSX
