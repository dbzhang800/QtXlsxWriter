/****************************************************************************
** Copyright (c) 2013 Debao Zhang <hello@debao.me>
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
#include "xlsxcontenttypes_p.h"
#include "xmlstreamwriter_p.h"
#include <QFile>
#include <QMapIterator>

namespace QXlsx {

ContentTypes::ContentTypes()
{
    m_package_prefix = "application/vnd.openxmlformats-package.";
    m_document_prefix = "application/vnd.openxmlformats-officedocument.";

    m_defaults.insert("rels", m_package_prefix + "relationships+xml");
    m_defaults.insert("xml", "application/xml");

    m_overrides.insert("/docProps/app.xml", m_document_prefix + "extended-properties+xml");
    m_overrides.insert("/docProps/core.xml", m_package_prefix + "core-properties+xml");
    m_overrides.insert("/xl/styles.xml", m_document_prefix + "spreadsheetml.styles+xml");
    m_overrides.insert("/xl/theme/theme1.xml", m_document_prefix + "theme+xml");
    m_overrides.insert("/xl/workbook.xml", m_document_prefix + "spreadsheetml.sheet.main+xml");
}

void ContentTypes::addDefault(const QString &key, const QString &value)
{
    m_defaults.insert(key, value);
}

void ContentTypes::addOverride(const QString &key, const QString &value)
{
    m_overrides.insert(key, value);
}

void ContentTypes::addWorksheetName(const QString &name)
{
    addOverride(QString("/xl/worksheets/%1.xml").arg(name), m_document_prefix + "spreadsheetml.worksheet+xml");
}

void ContentTypes::addChartsheetName(const QString &name)
{
    addOverride(QString("/xl/chartsheets/%1.xml").arg(name), m_document_prefix + "spreadsheetml.chartsheet+xml");
}

void ContentTypes::addChartName(const QString &name)
{
    addOverride(QString("/xl/charts/%1.xml").arg(name), m_document_prefix + "drawingml.chart+xml");
}

void ContentTypes::addCommentName(const QString &name)
{
    addOverride(QString("/xl/%1.xml").arg(name), m_document_prefix + "spreadsheetml.comments+xml");
}

void ContentTypes::addImageTypes(const QStringList &imageTypes)
{
    foreach (QString type, imageTypes)
        addOverride(type, "image/" + type);
}

void ContentTypes::addTableName(const QString &name)
{
    addOverride(QString("/xl/tables/%1.xml").arg(name), m_document_prefix + "spreadsheetml.table+xml");
}

void ContentTypes::addSharedString()
{
    addOverride("/xl/sharedStrings.xml", m_document_prefix + "spreadsheetml.sharedStrings+xml");
}

void ContentTypes::addVmlName()
{
    addOverride("vml", m_document_prefix + "vmlDrawing");
}

void ContentTypes::addCalcChain()
{
    addOverride("/xl/calcChain.xml", m_document_prefix + "spreadsheetml.calcChain+xml");
}

void ContentTypes::addVbaProject()
{
    //:TODO
    addOverride("bin", "application/vnd.ms-office.vbaProject");
}

void ContentTypes::saveToXmlFile(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument("1.0", true);
    writer.writeStartElement("Types");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

    {
    QMapIterator<QString, QString> it(m_defaults);
    while(it.hasNext()) {
        it.next();
        writer.writeStartElement("Default");
        writer.writeAttribute("Extension", it.key());
        writer.writeAttribute("ContentType", it.value());
        writer.writeEndElement();//Default
    }
    }

    {
    QMapIterator<QString, QString> it(m_overrides);
    while(it.hasNext()) {
        it.next();
        writer.writeStartElement("Override");
        writer.writeAttribute("PartName", it.key());
        writer.writeAttribute("ContentType", it.value());
        writer.writeEndElement(); //Override
    }
    }

    writer.writeEndElement();//Types
    writer.writeEndDocument();

}

} //namespace QXlsx
