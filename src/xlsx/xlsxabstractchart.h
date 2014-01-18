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

#ifndef QXLSX_XLSXABSTRACTCHART_H
#define QXLSX_XLSXABSTRACTCHART_H

#include "xlsxglobal.h"

#include <QString>

class QXmlStreamReader;
class QXmlStreamWriter;

QT_BEGIN_NAMESPACE_XLSX

class ChartFile;
class AbstractChartPrivate;
class CellRange;

class Q_XLSX_EXPORT AbstractChart
{
    Q_DECLARE_PRIVATE(AbstractChart)

public:
    AbstractChart();
    virtual ~AbstractChart();

    void addSeries(const CellRange &range, const QString &sheet=QString());

protected:
    friend class ChartFile;
    AbstractChart(AbstractChartPrivate *d);
    virtual bool loadXxxChartFromXml(QXmlStreamReader &reader) = 0;
    virtual void saveXxxChartToXml(QXmlStreamWriter &writer) const = 0;

    bool loadAxisFromXml(QXmlStreamReader &reader);
    void saveAxisToXml(QXmlStreamWriter &writer) const;

    bool loadLegendFromXml(QXmlStreamReader &reader);
    void saveLegendToXml(QXmlStreamWriter &writer) const;

    AbstractChartPrivate * d_ptr;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXABSTRACTCHART_H
