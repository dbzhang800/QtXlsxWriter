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

#include "xlsxpiechart.h"
#include "xlsxpiechart_p.h"

#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_XLSX

PieChartPrivate::PieChartPrivate(PieChart *chart)
    : AbstractChartPrivate(chart)
{

}

/*!
 * \class PieChart
 */

PieChart::PieChart()
    : AbstractChart(new PieChartPrivate(this))
{
}

bool PieChart::loadFromXml(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("pieChart"));

    Q_D(PieChart);

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("ser")) {
                d->loadXmlSer(reader);
            }
        } else if (reader.tokenType() == QXmlStreamReader::EndElement
                   && reader.name() == QLatin1String("pieChart")) {
            break;
        }
    }
    return true;
}

bool PieChartPrivate::loadXmlSer(QXmlStreamReader &reader)
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

void PieChart::saveToXml(QXmlStreamWriter &writer) const
{
    Q_D(const PieChart);

    writer.writeStartElement(QStringLiteral("c:pieChart"));
    writer.writeEmptyElement(QStringLiteral("c:varyColors"));
    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("1"));
    for (int i=0; i<d->seriesList.size(); ++i) {
        XlsxSeries *ser = d->seriesList[i].data();
        writer.writeStartElement(QStringLiteral("c:ser"));
        writer.writeEmptyElement(QStringLiteral("c:idx"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(i));
        writer.writeEmptyElement(QStringLiteral("c:order"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(i));
        writer.writeStartElement(QStringLiteral("c:val"));
        writer.writeStartElement(QStringLiteral("c:numRef"));
        writer.writeTextElement(QStringLiteral("c:f"), ser->numRef);
        writer.writeEndElement();//c:numRef
        writer.writeEndElement();//c:val
        writer.writeEndElement();//c:ser
    }

    writer.writeEndElement(); //pieChart
}

QT_END_NAMESPACE_XLSX
