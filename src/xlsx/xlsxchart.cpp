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

#include "xlsxchart_p.h"
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"
#include "xlsxutility_p.h"

#include <QIODevice>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_XLSX

ChartPrivate::ChartPrivate(Chart *q)
    :OOXmlFilePrivate(q), chartType(static_cast<Chart::ChartType>(0))
{

}

ChartPrivate::~ChartPrivate()
{

}

/*!
 * \class Chart
 *
 * Main class for the charts.
 */

Chart::Chart(Worksheet *parent)
    :OOXmlFile(new ChartPrivate(this))
{
    d_func()->sheet = parent;
}

Chart::~Chart()
{
}

void Chart::addSeries(const CellRange &range, Worksheet *sheet)
{
    Q_D(Chart);

    QString serRef = sheet ? sheet->sheetName() : d->sheet->sheetName();

    serRef += QLatin1String("!");
    serRef += xl_rowcol_to_cell(range.firstRow(), range.firstColumn(), true, true);
    serRef += QLatin1String(":");
    serRef += xl_rowcol_to_cell(range.lastRow(), range.lastColumn(), true, true);

    XlsxSeries *series = new XlsxSeries;
    series->numRef = serRef;

    d->seriesList.append(QSharedPointer<XlsxSeries>(series));
}

/*!
 * Set the type of the chart to \a type
 */
void Chart::setChartType(ChartType type)
{
    Q_D(Chart);
    d->chartType = type;
}

void Chart::setChartStyle(int id)
{
    Q_UNUSED(id)
    //!Todo
}

/*!
 * \internal
 */
void Chart::saveToXmlFile(QIODevice *device) const
{
    Q_D(const Chart);

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("c:chartSpace"));
    writer.writeAttribute(QStringLiteral("xmlns:c"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/chart"));
    writer.writeAttribute(QStringLiteral("xmlns:a"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    d->saveXmlChart(writer);

    writer.writeEndElement();//c:chartSpace
    writer.writeEndDocument();
}

/*!
 * \internal
 */
bool Chart::loadFromXmlFile(QIODevice *device)
{
    Q_D(Chart);

    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("chart")) {
                if (!d->loadXmlChart(reader))
                    return false;
            }
        }
    }
    return true;
}

bool ChartPrivate::loadXmlChart(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("chart"));

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("plotArea")) {
                if (!loadXmlPlotArea(reader))
                    return false;
            } else if (reader.name() == QLatin1String("legend")) {
                //!Todo
            }
        } else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                   reader.name() == QLatin1String("chart")) {
            break;
        }
    }
    return true;
}

bool ChartPrivate::loadXmlPlotArea(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("plotArea"));

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("layout")) {
                //!ToDo
            } else if (reader.name().endsWith(QLatin1String("Chart"))) {
                //For pieChart, barChart, ...
                loadXmlXxxChart(reader);
            } else if (reader.name().endsWith(QLatin1String("Ax"))) {
                //For valAx, catAx, serAx, dateAx
                loadXmlAxis(reader);
            }

        } else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                   reader.name() == QLatin1String("plotArea")) {
            break;
        }
    }
    return true;
}

bool ChartPrivate::loadXmlXxxChart(QXmlStreamReader &reader)
{
    QStringRef name = reader.name();
    if (name == QLatin1String("pieChart")) chartType = Chart::CT_Pie;
    else if (name == QLatin1String("pie3DChart")) chartType = Chart::CT_Pie3D;
    else if (name == QLatin1String("barChart")) chartType = Chart::CT_Bar;
    else if (name == QLatin1String("bar3DChart")) chartType = Chart::CT_Bar3D;

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("ser")) {
                loadXmlSer(reader);
            } else if (reader.name() == QLatin1String("axId")) {

            }
        } else if (reader.tokenType() == QXmlStreamReader::EndElement
                   && reader.name() == name) {
            break;
        }
    }
    return true;
}

bool ChartPrivate::loadXmlSer(QXmlStreamReader &reader)
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

void ChartPrivate::saveXmlChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QStringLiteral("c:chart"));
    writer.writeStartElement(QStringLiteral("c:plotArea"));
    saveXmlXxxChart(writer);
    saveXmlAxes(writer);
    writer.writeEndElement(); //plotArea

//    saveXmlLegend(writer);

    writer.writeEndElement(); //chart
}

void ChartPrivate::saveXmlXxxChart(QXmlStreamWriter &writer) const
{
    QString t;
    switch (chartType) {
    case Chart::CT_Pie: t = QStringLiteral("c:pieChart"); break;
    case Chart::CT_Pie3D: t = QStringLiteral("c:pie3DChart"); break;
    case Chart::CT_Bar: t = QStringLiteral("c:barChart"); break;
    case Chart::CT_Bar3D: t = QStringLiteral("c:bar3DChart"); break;
    default: break;
    }

    writer.writeStartElement(t); //pieChart, barChart, ...

    if (chartType==Chart::CT_Bar || chartType==Chart::CT_Bar3D) {
        writer.writeEmptyElement(QStringLiteral("c:barDir"));
        writer.writeAttribute(QStringLiteral("val"), QStringLiteral("col"));
    }

    if (chartType==Chart::CT_Pie || chartType==Chart::CT_Pie3D) {
        //Do the same behavior as Excel, Pie prefer varyColors
        writer.writeEmptyElement(QStringLiteral("c:varyColors"));
        writer.writeAttribute(QStringLiteral("val"), QStringLiteral("1"));
    }

    for (int i=0; i<seriesList.size(); ++i)
        saveXmlSer(writer, seriesList[i].data(), i);

    if (chartType == Chart::CT_Bar || chartType==Chart::CT_Bar3D) {
        if (axisList.isEmpty()) {
            const_cast<ChartPrivate*>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::T_Cat, XlsxAxis::Left)));
            const_cast<ChartPrivate*>(this)->axisList.append(QSharedPointer<XlsxAxis>(new XlsxAxis(XlsxAxis::T_Val, XlsxAxis::Bottom)));
        }

        Q_ASSERT(axisList.size()==2 || (axisList.size()==3 && chartType==Chart::CT_Bar3D));

        for (int i=0; i<axisList.size(); ++i) {
            writer.writeEmptyElement(QStringLiteral("c:axId"));
            writer.writeAttribute(QStringLiteral("val"), QString::number(i+1));
        }
    }

    writer.writeEndElement(); //pieChart, barChart, ...
}

void ChartPrivate::saveXmlSer(QXmlStreamWriter &writer, XlsxSeries *ser, int id) const
{
    writer.writeStartElement(QStringLiteral("c:ser"));
    writer.writeEmptyElement(QStringLiteral("c:idx"));
    writer.writeAttribute(QStringLiteral("val"), QString::number(id));
    writer.writeEmptyElement(QStringLiteral("c:order"));
    writer.writeAttribute(QStringLiteral("val"), QString::number(id));
    writer.writeStartElement(QStringLiteral("c:val"));
    writer.writeStartElement(QStringLiteral("c:numRef"));
    writer.writeTextElement(QStringLiteral("c:f"), ser->numRef);
    writer.writeEndElement();//c:numRef
    writer.writeEndElement();//c:val
    writer.writeEndElement();//c:ser
}

bool ChartPrivate::loadXmlAxis(QXmlStreamReader &reader)
{
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

    axisList.append(QSharedPointer<XlsxAxis>(axis));

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

void ChartPrivate::saveXmlAxes(QXmlStreamWriter &writer) const
{
    for (int i=0; i<axisList.size(); ++i) {
        XlsxAxis *axis = axisList[i].data();
        QString name;
        switch (axis->type) {
        case XlsxAxis::T_Cat: name = QStringLiteral("c:catAx"); break;
        case XlsxAxis::T_Val: name = QStringLiteral("c:valAx"); break;
        case XlsxAxis::T_Ser: name = QStringLiteral("c:serAx"); break;
        case XlsxAxis::T_Date: name = QStringLiteral("c:dateAx"); break;
        default: break;
        }

        QString pos;
        switch (axis->axisPos) {
        case XlsxAxis::Top: pos = QStringLiteral("t"); break;
        case XlsxAxis::Bottom: pos = QStringLiteral("b"); break;
        case XlsxAxis::Left: pos = QStringLiteral("l"); break;
        case XlsxAxis::Right: pos = QStringLiteral("r"); break;
        default: break;
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

QT_END_NAMESPACE_XLSX
