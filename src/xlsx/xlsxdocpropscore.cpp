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
#include "xlsxdocpropscore_p.h"
#include "xlsxxmlwriter_p.h"
#include "xlsxxmlreader_p.h"

#include <QDir>
#include <QFile>
#include <QDateTime>
#include <QDebug>
#include <QBuffer>

namespace QXlsx {

DocPropsCore::DocPropsCore()
{
}

bool DocPropsCore::setProperty(const QString &name, const QString &value)
{
    static QStringList validKeys;
    if (validKeys.isEmpty()) {
        validKeys << QStringLiteral("title") << QStringLiteral("subject")
                  << QStringLiteral("keywords") << QStringLiteral("description")
                  << QStringLiteral("category") << QStringLiteral("status")
                  << QStringLiteral("created") << QStringLiteral("creator");
    }

    if (!validKeys.contains(name))
        return false;

    if (value.isEmpty())
        m_properties.remove(name);
    else
        m_properties[name] = value;

    return true;
}

QString DocPropsCore::property(const QString &name) const
{
    if (m_properties.contains(name))
        return m_properties[name];

    return QString();
}

QStringList DocPropsCore::propertyNames() const
{
    return m_properties.keys();
}

void DocPropsCore::saveToXmlFile(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("cp:coreProperties"));
    writer.writeAttribute(QStringLiteral("xmlns:cp"), QStringLiteral("http://schemas.openxmlformats.org/package/2006/metadata/core-properties"));
    writer.writeAttribute(QStringLiteral("xmlns:dc"), QStringLiteral("http://purl.org/dc/elements/1.1/"));
    writer.writeAttribute(QStringLiteral("xmlns:dcterms"), QStringLiteral("http://purl.org/dc/terms/"));
    writer.writeAttribute(QStringLiteral("xmlns:dcmitype"), QStringLiteral("http://purl.org/dc/dcmitype/"));
    writer.writeAttribute(QStringLiteral("xmlns:xsi"), QStringLiteral("http://www.w3.org/2001/XMLSchema-instance"));

    if (m_properties.contains(QStringLiteral("title")))
        writer.writeTextElement(QStringLiteral("dc:title"), m_properties[QStringLiteral("title")]);

    if (m_properties.contains(QStringLiteral("subject")))
        writer.writeTextElement(QStringLiteral("dc:subject"), m_properties[QStringLiteral("subject")]);

    writer.writeTextElement(QStringLiteral("dc:creator"), m_properties.contains(QStringLiteral("creator")) ? m_properties[QStringLiteral("creator")] : QStringLiteral("Qt Xlsx Library"));

    if (m_properties.contains(QStringLiteral("keywords")))
        writer.writeTextElement(QStringLiteral("cp:keywords"), m_properties[QStringLiteral("keywords")]);

    if (m_properties.contains(QStringLiteral("description")))
        writer.writeTextElement(QStringLiteral("dc:description"), m_properties[QStringLiteral("description")]);

    writer.writeTextElement(QStringLiteral("cp:lastModifiedBy"), m_properties.contains(QStringLiteral("creator")) ? m_properties[QStringLiteral("creator")] : QStringLiteral("Qt Xlsx Library"));

    writer.writeStartElement(QStringLiteral("dcterms:created"));
    writer.writeAttribute(QStringLiteral("xsi:type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(m_properties.contains(QStringLiteral("created")) ? m_properties[QStringLiteral("created")] : QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writer.writeStartElement(QStringLiteral("dcterms:modified"));
    writer.writeAttribute(QStringLiteral("xsi:type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    if (m_properties.contains(QStringLiteral("category")))
        writer.writeTextElement(QStringLiteral("cp:category"), m_properties[QStringLiteral("category")]);

    if (m_properties.contains(QStringLiteral("status")))
        writer.writeTextElement(QStringLiteral("cp:contentStatus"), m_properties[QStringLiteral("status")]);

    writer.writeEndElement(); //cp:coreProperties
    writer.writeEndDocument();
}

QByteArray DocPropsCore::saveToXmlData()
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);

    return data;
}

DocPropsCore DocPropsCore::loadFromXmlFile(QIODevice *device)
{
    DocPropsCore props;
    XmlStreamReader reader(device);
    while(!reader.atEnd()) {
         QXmlStreamReader::TokenType token = reader.readNext();
         if (token == QXmlStreamReader::StartElement) {
             if (reader.qualifiedName() == QLatin1String("cp:coreProperties"))
                 continue;

             QString text = reader.readElementText();
             if (reader.qualifiedName() == QStringLiteral("dc:subject")) {
                 props.setProperty(QStringLiteral("subject"), text);
             } else if (reader.qualifiedName() == QStringLiteral("dc:title")) {
                 props.setProperty(QStringLiteral("title"), text);
             } else if (reader.qualifiedName() == QStringLiteral("dc:creator")) {
                 props.setProperty(QStringLiteral("creator"), text);
             } else if (reader.qualifiedName() == QStringLiteral("dc:description")) {
                 props.setProperty(QStringLiteral("description"), text);
             } else if (reader.qualifiedName() == QStringLiteral("cp:keywords")) {
                 props.setProperty(QStringLiteral("keywords"), text);
             } else if (reader.qualifiedName() == QStringLiteral("dcterms:created")) {
                 props.setProperty(QStringLiteral("created"), text);
             } else if (reader.qualifiedName() == QStringLiteral("cp:category")) {
                 props.setProperty(QStringLiteral("category"), text);
             } else if (reader.qualifiedName() == QStringLiteral("cp:contentStatus")) {
                 props.setProperty(QStringLiteral("status"), text);
             }
         }

         if (reader.hasError()) {
             qDebug()<<"Error when read doc props core file.";

         }
    }
    return props;
}

DocPropsCore DocPropsCore::loadFromXmlData(const QByteArray &data)
{
    QBuffer buffer;
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);
    return loadFromXmlFile(&buffer);
}


} //namespace
