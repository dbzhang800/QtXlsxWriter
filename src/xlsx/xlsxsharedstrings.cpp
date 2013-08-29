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
#include "xlsxsharedstrings_p.h"
#include "xmlstreamwriter_p.h"
#include <QDir>
#include <QFile>
#include <QRegularExpression>
#include <QDebug>

namespace QXlsx {

SharedStrings::SharedStrings(QObject *parent) :
    QObject(parent)
{
    m_stringCount = 0;
}

int SharedStrings::count() const
{
    return m_stringCount;
}

int SharedStrings::addSharedString(const QString &string)
{
    m_stringCount += 1;

    if (m_stringTable.contains(string))
        return m_stringTable[string];

    int index = m_stringTable.size();
    m_stringTable[string] = index;
    m_stringList.append(string);
    return index;
}

int SharedStrings::getSharedStringIndex(const QString &string) const
{
    return m_stringTable[string];
}

QString SharedStrings::getSharedString(int index) const
{
    return m_stringList[index];
}

QStringList SharedStrings::getSharedStrings() const
{
    return m_stringList;
}

void SharedStrings::saveToXmlFile(QIODevice *device) const
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument("1.0", true);
    writer.writeStartElement("sst");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("count", QString::number(m_stringCount));
    writer.writeAttribute("uniqueCount", QString::number(m_stringTable.size()));

    foreach (QString string, m_stringList) {
        writer.writeStartElement("si");
        if (string.contains(QRegularExpression("^<r>")) || string.contains(QRegularExpression("</r>$"))) {
            //Rich text string,
//            writer.writeCharacters(string);
        } else {
            writer.writeStartElement("t");
            if (string.contains(QRegularExpression("^\\s")) || string.contains(QRegularExpression("\\s$")))
                writer.writeAttribute("xml:space", "preserve");
            writer.writeCharacters(string);
            writer.writeEndElement();//t
        }
        writer.writeEndElement();//si
    }

    writer.writeEndElement(); //sst
    writer.writeEndDocument();
}

} //namespace
