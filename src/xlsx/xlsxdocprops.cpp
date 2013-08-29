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
#include "xmlstreamwriter_p.h"

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

    writer.writeStartDocument("1.0", true);
    writer.writeStartElement("Properties");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
    writer.writeAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
    writer.writeTextElement("Application", "Microsoft Excel");
    writer.writeTextElement("DocSecurity", "0");
    writer.writeTextElement("ScaleCrop", "false");

    writer.writeStartElement("HeadingPairs");
    writer.writeStartElement("vt:vector");
    writer.writeAttribute("size", QString::number(m_headingPairsList.size()*2));
    writer.writeAttribute("baseType", "variant");
    typedef QPair<QString,int> PairType; //Make foreach happy
    foreach (PairType pair, m_headingPairsList) {
        writer.writeStartElement("vt:variant");
        writer.writeTextElement("vt:lpstr", pair.first);
        writer.writeEndElement(); //vt:variant
        writer.writeStartElement("vt:variant");
        writer.writeTextElement("vt:i4", QString::number(pair.second));
        writer.writeEndElement(); //vt:variant
    }
    writer.writeEndElement();//vt:vector
    writer.writeEndElement();//HeadingPairs

    writer.writeStartElement("TitlesOfParts");
    writer.writeStartElement("vt:vector");
    writer.writeAttribute("size", QString::number(m_titlesOfPartsList.size()));
    writer.writeAttribute("baseType", "lpstr");
    foreach (QString title, m_titlesOfPartsList)
        writer.writeTextElement("vt:lpstr", title);
    writer.writeEndElement();//vt:vector
    writer.writeEndElement();//TitlesOfParts

    writer.writeTextElement("Company", "");
    writer.writeTextElement("LinksUpToDate", "false");
    writer.writeTextElement("SharedDoc", "false");
    writer.writeTextElement("HyperlinksChanged", "false");
    writer.writeTextElement("AppVersion", "12.0000");
    writer.writeEndElement(); //Properties
    writer.writeEndDocument();
}

void DocProps::saveToXmlFile_Core(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument("1.0", true);
    writer.writeStartElement("cp:coreProperties");
    writer.writeAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    writer.writeAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
    writer.writeAttribute("xmlns:dcterms", "http://purl.org/dc/terms/");
    writer.writeAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
    writer.writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
    writer.writeTextElement("dc:title", "");
    writer.writeTextElement("dc:subject", "");
    writer.writeTextElement("dc:creator", "QXlsxWriter");
    writer.writeTextElement("cp:keywords", "");
    writer.writeTextElement("dc:description", "");
    writer.writeTextElement("cp:lastModifiedBy", "");

    writer.writeStartElement("dcterms:created");
    writer.writeAttribute("xsi:type", "dcterms:W3CDTF");
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writer.writeStartElement("dcterms:modified");
    writer.writeAttribute("xsi:type", "dcterms:W3CDTF");
    writer.writeCharacters(QDateTime::currentDateTime().toString(Qt::ISODate));
    writer.writeEndElement();//dcterms:created

    writer.writeTextElement("cp:category", "");
    writer.writeTextElement("cp:contentStatus", "");
    writer.writeEndElement(); //cp:coreProperties
    writer.writeEndDocument();
}

} //namespace
