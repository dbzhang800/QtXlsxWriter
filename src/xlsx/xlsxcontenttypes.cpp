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
#include "xlsxcontenttypes_p.h"
#include <QXmlStreamWriter>
#include <QFile>
#include <QMapIterator>
#include <QBuffer>

namespace QXlsx {

ContentTypes::ContentTypes()
{
    m_package_prefix = QStringLiteral("application/vnd.openxmlformats-package.");
    m_document_prefix = QStringLiteral("application/vnd.openxmlformats-officedocument.");

    m_defaults.insert(QStringLiteral("rels"), m_package_prefix + QStringLiteral("relationships+xml"));
    m_defaults.insert(QStringLiteral("xml"), QStringLiteral("application/xml"));

    m_overrides.insert(QStringLiteral("/docProps/app.xml"), m_document_prefix + QStringLiteral("extended-properties+xml"));
    m_overrides.insert(QStringLiteral("/docProps/core.xml"), m_package_prefix + QStringLiteral("core-properties+xml"));
    m_overrides.insert(QStringLiteral("/xl/styles.xml"), m_document_prefix + QStringLiteral("spreadsheetml.styles+xml"));
    m_overrides.insert(QStringLiteral("/xl/theme/theme1.xml"), m_document_prefix + QStringLiteral("theme+xml"));
    m_overrides.insert(QStringLiteral("/xl/workbook.xml"), m_document_prefix + QStringLiteral("spreadsheetml.sheet.main+xml"));
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
    addOverride(QStringLiteral("/xl/worksheets/%1.xml").arg(name), m_document_prefix + QStringLiteral("spreadsheetml.worksheet+xml"));
}

void ContentTypes::addChartsheetName(const QString &name)
{
    addOverride(QStringLiteral("/xl/chartsheets/%1.xml").arg(name), m_document_prefix + QStringLiteral("spreadsheetml.chartsheet+xml"));
}

void ContentTypes::addDrawingName(const QString &name)
{
    addOverride(QStringLiteral("/xl/drawings/%1.xml").arg(name), m_document_prefix + QStringLiteral("drawing+xml"));
}

void ContentTypes::addChartName(const QString &name)
{
    addOverride(QStringLiteral("/xl/charts/%1.xml").arg(name), m_document_prefix + QStringLiteral("drawingml.chart+xml"));
}

void ContentTypes::addCommentName(const QString &name)
{
    addOverride(QStringLiteral("/xl/%1.xml").arg(name), m_document_prefix + QStringLiteral("spreadsheetml.comments+xml"));
}

void ContentTypes::addImageTypes(const QStringList &imageTypes)
{
    foreach (QString type, imageTypes)
        addDefault(type, QStringLiteral("image/") + type);
}

void ContentTypes::addTableName(const QString &name)
{
    addOverride(QStringLiteral("/xl/tables/%1.xml").arg(name), m_document_prefix + QStringLiteral("spreadsheetml.table+xml"));
}

void ContentTypes::addSharedString()
{
    addOverride(QStringLiteral("/xl/sharedStrings.xml"), m_document_prefix + QStringLiteral("spreadsheetml.sharedStrings+xml"));
}

void ContentTypes::addVmlName()
{
    addOverride(QStringLiteral("vml"), m_document_prefix + QStringLiteral("vmlDrawing"));
}

void ContentTypes::addCalcChain()
{
    addOverride(QStringLiteral("/xl/calcChain.xml"), m_document_prefix + QStringLiteral("spreadsheetml.calcChain+xml"));
}

void ContentTypes::addVbaProject()
{
    //:TODO
    addOverride(QStringLiteral("bin"), QStringLiteral("application/vnd.ms-office.vbaProject"));
}

QByteArray ContentTypes::saveToXmlData() const
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);
    return data;
}

void ContentTypes::saveToXmlFile(QIODevice *device) const
{
    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("Types"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/package/2006/content-types"));

    {
    QMapIterator<QString, QString> it(m_defaults);
    while (it.hasNext()) {
        it.next();
        writer.writeStartElement(QStringLiteral("Default"));
        writer.writeAttribute(QStringLiteral("Extension"), it.key());
        writer.writeAttribute(QStringLiteral("ContentType"), it.value());
        writer.writeEndElement();//Default
    }
    }

    {
    QMapIterator<QString, QString> it(m_overrides);
    while (it.hasNext()) {
        it.next();
        writer.writeStartElement(QStringLiteral("Override"));
        writer.writeAttribute(QStringLiteral("PartName"), it.key());
        writer.writeAttribute(QStringLiteral("ContentType"), it.value());
        writer.writeEndElement(); //Override
    }
    }

    writer.writeEndElement();//Types
    writer.writeEndDocument();

}

} //namespace QXlsx
