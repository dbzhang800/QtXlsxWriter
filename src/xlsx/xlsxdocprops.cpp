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
#include "xlsxdocprops_p.h"
#include "xlsxxmlwriter_p.h"

#include <QDir>
#include <QFile>
#include <QDateTime>
namespace QXlsx {

DocProps::DocProps(QObject *parent) :
    QObject(parent)
{
}

void DocProps::addPartTitle(const QString &title)
{
    m_titlesOfPartsList.append(title);
}

void DocProps::addHeadingPair(const QString &name, int value)
{
    m_headingPairsList.append(qMakePair(name, value));
}


void DocProps::saveToXmlFile_App(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("Properties"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"));
    writer.writeAttribute(QStringLiteral("xmlns:vt"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"));
    writer.writeTextElement(QStringLiteral("Application"), QStringLiteral("Microsoft Excel"));
    writer.writeTextElement(QStringLiteral("DocSecurity"), QStringLiteral("0"));
    writer.writeTextElement(QStringLiteral("ScaleCrop"), QStringLiteral("false"));

    writer.writeStartElement(QStringLiteral("HeadingPairs"));
    writer.writeStartElement(QStringLiteral("vt:vector"));
    writer.writeAttribute(QStringLiteral("size"), QString::number(m_headingPairsList.size()*2));
    writer.writeAttribute(QStringLiteral("baseType"), QStringLiteral("variant"));
    typedef QPair<QString,int> PairType; //Make foreach happy
    foreach (PairType pair, m_headingPairsList) {
        writer.writeStartElement(QStringLiteral("vt:variant"));
        writer.writeTextElement(QStringLiteral("vt:lpstr"), pair.first);
        writer.writeEndElement(); //vt:variant
        writer.writeStartElement(QStringLiteral("vt:variant"));
        writer.writeTextElement(QStringLiteral("vt:i4"), QString::number(pair.second));
        writer.writeEndElement(); //vt:variant
    }
    writer.writeEndElement();//vt:vector
    writer.writeEndElement();//HeadingPairs

    writer.writeStartElement(QStringLiteral("TitlesOfParts"));
    writer.writeStartElement(QStringLiteral("vt:vector"));
    writer.writeAttribute(QStringLiteral("size"), QString::number(m_titlesOfPartsList.size()));
    writer.writeAttribute(QStringLiteral("baseType"), QStringLiteral("lpstr"));
    foreach (QString title, m_titlesOfPartsList)
        writer.writeTextElement(QStringLiteral("vt:lpstr"), title);
    writer.writeEndElement();//vt:vector
    writer.writeEndElement();//TitlesOfParts

    writer.writeTextElement(QStringLiteral("Company"), QStringLiteral(""));
    writer.writeTextElement(QStringLiteral("LinksUpToDate"), QStringLiteral("false"));
    writer.writeTextElement(QStringLiteral("SharedDoc"), QStringLiteral("false"));
    writer.writeTextElement(QStringLiteral("HyperlinksChanged"), QStringLiteral("false"));
    writer.writeTextElement(QStringLiteral("AppVersion"), QStringLiteral("12.0000"));
    writer.writeEndElement(); //Properties
    writer.writeEndDocument();
}

void DocProps::saveToXmlFile_Core(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("cp:coreProperties"));
    writer.writeAttribute(QStringLiteral("xmlns:cp"), QStringLiteral("http://schemas.openxmlformats.org/package/2006/metadata/core-properties"));
    writer.writeAttribute(QStringLiteral("xmlns:dc"), QStringLiteral("http://purl.org/dc/elements/1.1/"));
    writer.writeAttribute(QStringLiteral("xmlns:dcterms"), QStringLiteral("http://purl.org/dc/terms/"));
    writer.writeAttribute(QStringLiteral("xmlns:dcmitype"), QStringLiteral("http://purl.org/dc/dcmitype/"));
    writer.writeAttribute(QStringLiteral("xmlns:xsi"), QStringLiteral("http://www.w3.org/2001/XMLSchema-instance"));
    writer.writeTextElement(QStringLiteral("dc:title"), QStringLiteral(""));
    writer.writeTextElement(QStringLiteral("dc:subject"), QStringLiteral(""));
    writer.writeTextElement(QStringLiteral("dc:creator"), QStringLiteral("QXlsxWriter"));
    writer.writeTextElement(QStringLiteral("cp:keywords"), QStringLiteral(""));
    writer.writeTextElement(QStringLiteral("dc:description"), QStringLiteral(""));
    writer.writeTextElement(QStringLiteral("cp:lastModifiedBy"), QStringLiteral(""));

    writer.writeStartElement(QStringLiteral("dcterms:created"));
    writer.writeAttribute(QStringLiteral("xsi:type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writer.writeStartElement(QStringLiteral("dcterms:modified"));
    writer.writeAttribute(QStringLiteral("xsi:type"), QStringLiteral("dcterms:W3CDTF"));
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writer.writeTextElement(QStringLiteral("cp:category"), QStringLiteral(""));
    writer.writeTextElement(QStringLiteral("cp:contentStatus"), QStringLiteral(""));
    writer.writeEndElement(); //cp:coreProperties
    writer.writeEndDocument();
}

} //namespace
