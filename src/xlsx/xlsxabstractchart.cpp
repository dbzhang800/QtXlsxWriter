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
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_XLSX

AbstractChartPrivate::AbstractChartPrivate(AbstractChart *chart)
    :q_ptr(chart)
{

}

AbstractChartPrivate::~AbstractChartPrivate()
{

}

bool AbstractChartPrivate::loadXmlSer(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("ser"));

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("f")) {
                XlsxSeries *series = new XlsxSeries;
                series->numRef = reader.readElementText();
                seriesList.append(QSharedPointer<XlsxSeries>(series));
            }
        } else if (reader.tokenType() == QXmlStreamReader::EndElement
                   && reader.name() == QLatin1String("ser")) {
            break;
        }
    }

    return true;
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

bool AbstractChart::loadAxisFromXml(QXmlStreamReader &reader)
{
    Q_D(AbstractChart);
    Q_ASSERT(reader.name().endsWith(QLatin1String("Ax")));
    QString name = reader.name().toString();

    XlsxAxis *axis = new XlsxAxis;
    if (name == QLatin1String("valAx"))
        axis->type = XlsxAxis::T_Val;
    else if (name == QLatin1String("catAx"))
        axis->type = XlsxAxis::T_Cat;
    else if (name == QLatin1String("serAx"))
        axis->type = XlsxAxis::T_Ser;
    else
        axis->type = XlsxAxis::T_Date;

    d->axisList.append(QSharedPointer<XlsxAxis>(axis));

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("axPos")) {
                QXmlStreamAttributes attrs = reader.attributes();
                QStringRef pos = attrs.value(QLatin1String("val"));
                if (pos==QLatin1String("l"))
                    axis->axisPos = XlsxAxis::Left;
                else if (pos==QLatin1String("r"))
                    axis->axisPos = XlsxAxis::Right;
                else if (pos==QLatin1String("b"))
                    axis->axisPos = XlsxAxis::Bottom;
                else
                    axis->axisPos = XlsxAxis::Top;
            }
        } else if (reader.tokenType() == QXmlStreamReader::EndElement
                   && reader.name() == name) {
            break;
        }
    }

    return true;
}

void AbstractChart::saveAxisToXml(QXmlStreamWriter &writer) const
{
    Q_D(const AbstractChart);

    for (int i=0; i<d->axisList.size(); ++i) {
        XlsxAxis *axis = d->axisList[i].data();
        QString name;
        switch (axis->type) {
        case XlsxAxis::T_Cat:
            name = QStringLiteral("c:catAx");
            break;
        case XlsxAxis::T_Val:
            name = QStringLiteral("c:valAx");
            break;
        case XlsxAxis::T_Ser:
            name = QStringLiteral("c:serAx");
            break;
        case XlsxAxis::T_Date:
            name = QStringLiteral("c:dateAx");
            break;
        default:
            break;
        }

        QString pos;
        switch (axis->axisPos) {
        case XlsxAxis::Top:
            pos = QStringLiteral("t");
            break;
        case XlsxAxis::Bottom:
            pos = QStringLiteral("b");
            break;
        case XlsxAxis::Left:
            pos = QStringLiteral("l");
            break;
        case XlsxAxis::Right:
            pos = QStringLiteral("r");
            break;
        default:
            break;
        }

        writer.writeStartElement(name);
        writer.writeEmptyElement(QStringLiteral("c:axId"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(i+1));

        writer.writeStartElement(QStringLiteral("c:scaling"));
        writer.writeEmptyElement(QStringLiteral("c:orientation"));
        writer.writeAttribute(QStringLiteral("val"), QStringLiteral("minMax"));
        writer.writeEndElement();//c:scaling

        writer.writeEmptyElement(QStringLiteral("c:axPos"));
        writer.writeAttribute(QStringLiteral("val"), pos);

        writer.writeEmptyElement(QStringLiteral("c:crossAx"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(i%2==0 ? i+2 : i));

        writer.writeEndElement();//name
    }
}

bool AbstractChart::loadLegendFromXml(QXmlStreamReader &reader)
{

    return false;
}

void AbstractChart::saveLegendToXml(QXmlStreamWriter &writer) const
{

}

QT_END_NAMESPACE_XLSX
