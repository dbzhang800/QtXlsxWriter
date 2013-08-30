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

#include <QDir>
#include <QFile>
#include <QDateTime>
#include <QVariant>
namespace QXlsx {

DocPropsCore::DocPropsCore(QObject *parent) :
    QObject(parent)
{
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
    if (property("title").isValid())
        writer.writeTextElement(QStringLiteral("dc:title"), property("title").toString());
    if (property("subject").isValid())
        writer.writeTextElement(QStringLiteral("dc:subject"), property("subject").toString());
    writer.writeTextElement(QStringLiteral("dc:creator"), property("creator").isValid() ? property("creator").toString() : QStringLiteral("Qt Xlsx Library"));

    if (property("keywords").isValid())
        writer.writeTextElement(QStringLiteral("cp:keywords"), property("keywords").toString());
    if (property("description").isValid())
        writer.writeTextElement(QStringLiteral("dc:description"), property("description").toString());
    writer.writeTextElement(QStringLiteral("cp:lastModifiedBy"), property("creator").isValid() ? property("creator").toString() : QStringLiteral("Qt Xlsx Library"));

    writer.writeStartElement(QStringLiteral("dcterms:created"));
    writer.writeAttribute(QStringLiteral("xsi:type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writer.writeStartElement(QStringLiteral("dcterms:modified"));
    writer.writeAttribute(QStringLiteral("xsi:type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    if (property("category").isValid())
        writer.writeTextElement(QStringLiteral("cp:category"), property("category").toString());
    if (property("status").isValid())
        writer.writeTextElement(QStringLiteral("cp:contentStatus"), property("status").toString());
    writer.writeEndElement(); //cp:coreProperties
    writer.writeEndDocument();
}

} //namespace
