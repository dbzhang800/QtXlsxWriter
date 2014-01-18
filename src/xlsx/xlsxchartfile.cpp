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

#include "xlsxchartfile_p.h"
#include "xlsxabstractchart.h"
#include "xlsxabstractchart_p.h"
#include "xlsxpiechart.h"
#include "xlsxbarchart.h"
#include <QIODevice>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

namespace QXlsx {

ChartFile::ChartFile()
{
}

ChartFile::~ChartFile()
{
    if (m_chart)
        delete m_chart;
}

AbstractChart* ChartFile::chart() const
{
    return m_chart;
}

void ChartFile::setChart(AbstractChart *chart)
{
    m_chart = chart;
    chart->d_func()->cf = this;
}

void ChartFile::saveToXmlFile(QIODevice *device) const
{
    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("c:chartSpace"));
    writer.writeAttribute(QStringLiteral("xmlns:c"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/chart"));
    writer.writeAttribute(QStringLiteral("xmlns:a"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    saveXmlChart(writer);

    writer.writeEndElement();//c:chartSpace
    writer.writeEndDocument();
}

bool ChartFile::loadFromXmlFile(QIODevice *device)
{
    QXmlStreamReader reader(device);
    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("chart")) {
                if (!loadXmlChart(reader))
                    return false;
            }
        }
    }
    return true;
}

bool ChartFile::loadXmlChart(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("chart"));

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("plotArea")) {
                if (!loadXmlPlotArea(reader))
                    return false;
            } else if (reader.name() == QLatin1String("legend")) {
                m_chart->loadLegendFromXml(reader);
            }
        } else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                   reader.name() == QLatin1String("chart")) {
            break;
        }
    }
    return true;
}

bool ChartFile::loadXmlPlotArea(QXmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("plotArea"));

    while (!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("layout")) {
                //...
            } else if (reader.name().endsWith(QLatin1String("Chart"))) {
                //!Todo: each plotArea can have more than one xxxChart
                AbstractChart *chart = 0;
                if (reader.name() == QLatin1String("pieChart")) {
                    chart = new PieChart;
                } else if (reader.name() == QLatin1String("barChart")) {
                    chart = new BarChart;
                } else {
                    //Not support
                    return false;
                }
                if (chart) {
                    chart->loadXxxChartFromXml(reader);
                    setChart(chart);
                }
            } else if (reader.name().endsWith(QLatin1String("Ax"))) {
                //For valAx, catAx, serAx, dateAx
                m_chart->loadAxisFromXml(reader);
            }

        } else if (reader.tokenType() == QXmlStreamReader::EndElement &&
                   reader.name() == QLatin1String("plotArea")) {
            break;
        }
    }
    return true;
}

bool ChartFile::saveXmlChart(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QStringLiteral("c:chart"));
    writer.writeStartElement(QStringLiteral("c:plotArea"));
    m_chart->saveXxxChartToXml(writer);
    m_chart->saveAxisToXml(writer);
    writer.writeEndElement(); //plotArea

    m_chart->saveLegendToXml(writer);

    writer.writeEndElement(); //chart
    return true;
}

} // namespace QXlsx
