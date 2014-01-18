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

#ifndef XLSXABSTRACTCHART_P_H
#define XLSXABSTRACTCHART_P_H

//
//  W A R N I N G
//  -------------
//
// This file is not part of the Qt Xlsx API.  It exists for the convenience
// of the Qt Xlsx.  This header file may change from
// version to version without notice, or even be removed.
//
// We mean it.
//

#include "xlsxabstractchart.h"
#include <QList>
#include <QSharedPointer>

QT_BEGIN_NAMESPACE_XLSX

class XlsxSeries
{
public:

    QString numRef;
};

class XlsxAxis
{
public:
    enum Type
    {
        T_Cat,
        T_Val,
        T_Date,
        T_Ser
    };

    enum Pos
    {
        Left,
        Right,
        Top,
        Bottom
    };

    XlsxAxis(){}

    XlsxAxis(Type t, Pos p)
        :type(t), axisPos(p)
    {
        axisId = -1;
    }

    int axisId;
    Type type;
    Pos axisPos; //l,r,b,t
};

class AbstractChartPrivate
{
    Q_DECLARE_PUBLIC(AbstractChart)
public:
    AbstractChartPrivate(AbstractChart *chart);
    virtual ~AbstractChartPrivate();

    bool loadXmlSer(QXmlStreamReader &reader);

    QList<QSharedPointer<XlsxSeries> > seriesList;
    QList<QSharedPointer<XlsxAxis>> axisList;
    ChartFile *cf;
    AbstractChart *q_ptr;
};

QT_END_NAMESPACE_XLSX
#endif // XLSXABSTRACTCHART_P_H
