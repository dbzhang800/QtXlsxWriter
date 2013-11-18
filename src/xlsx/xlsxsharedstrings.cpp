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
#include "xlsxrichstring_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxxmlwriter_p.h"
#include "xlsxxmlreader_p.h"
#include <QDir>
#include <QFile>
#include <QRegularExpression>
#include <QDebug>
#include <QBuffer>

namespace QXlsx {

SharedStrings::SharedStrings()
{
    m_stringCount = 0;
}

int SharedStrings::count() const
{
    return m_stringCount;
}

int SharedStrings::addSharedString(const QString &string)
{
    return addSharedString(RichString(string));
}

int SharedStrings::addSharedString(const RichString &string)
{
    m_stringCount += 1;

    if (m_stringTable.contains(string)) {
        XlsxSharedStringInfo &item = m_stringTable[string];
        item.count += 1;
        return item.index;
    }

    int index = m_stringTable.size();
    m_stringTable[string] = XlsxSharedStringInfo(index);
    m_stringList.append(string);
    return index;
}

void SharedStrings::incRefByStringIndex(int idx)
{
    if (idx <0 || idx >= m_stringList.size()) {
        qDebug("SharedStrings: invlid index");
        return;
    }

    addSharedString(m_stringList[idx]);
}

void SharedStrings::removeSharedString(const QString &string)
{
    removeSharedString(RichString(string));
}

void SharedStrings::removeSharedString(const RichString &string)
{
    if (!m_stringTable.contains(string))
        return;

    m_stringCount -= 1;

    XlsxSharedStringInfo &item = m_stringTable[string];
    item.count -= 1;

    if (item.count <= 0) {
        for (int i=item.index+1; i<m_stringList.size(); ++i)
            m_stringTable[m_stringList[i]].index -= 1;

        m_stringList.removeAt(item.index);
        m_stringTable.remove(string);
    }
}

int SharedStrings::getSharedStringIndex(const QString &string) const
{
    return getSharedStringIndex(RichString(string));
}

int SharedStrings::getSharedStringIndex(const RichString &string) const
{
    if (m_stringTable.contains(string))
        return m_stringTable[string].index;
    return -1;
}

RichString SharedStrings::getSharedString(int index) const
{
    if (index < m_stringList.count() && index >= 0)
        return m_stringList[index];
    return RichString();
}

QList<RichString> SharedStrings::getSharedStrings() const
{
    return m_stringList;
}

void SharedStrings::saveToXmlFile(QIODevice *device) const
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("sst"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_stringCount));
    writer.writeAttribute(QStringLiteral("uniqueCount"), QString::number(m_stringTable.size()));

    foreach (RichString string, m_stringList) {
        writer.writeStartElement(QStringLiteral("si"));
        if (string.isRichString()) {
            //Rich text string
            for (int i=0; i<string.fragmentCount(); ++i) {
                if (string.fragmentFormat(i)) {
                    writer.writeStartElement(QStringLiteral("rPr"));
                    //:Todo
                    writer.writeEndElement();// rPr
                }
                writer.writeStartElement(QStringLiteral("t"));
                writer.writeAttribute(QStringLiteral("xml:space"), QStringLiteral("preserve"));
                writer.writeCharacters(string.fragmentText(i));
                writer.writeEndElement();// t
            }
        } else {
            writer.writeStartElement(QStringLiteral("t"));
            QString pString = string.toPlainString();
            if (pString.contains(QRegularExpression(QStringLiteral("^\\s")))
                    || pString.contains(QRegularExpression(QStringLiteral("\\s$")))) {
                writer.writeAttribute(QStringLiteral("xml:space"), QStringLiteral("preserve"));
            }
            writer.writeCharacters(pString);
            writer.writeEndElement();//t
        }
        writer.writeEndElement();//si
    }

    writer.writeEndElement(); //sst
    writer.writeEndDocument();
}

QByteArray SharedStrings::saveToXmlData() const
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);

    return data;
}

void SharedStrings::readString(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("si"));

    RichString richString;

    while (!(reader.name() == QLatin1String("si") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("r"))
                readRichStringPart(reader, richString);
            else if (reader.name() == QLatin1String("t"))
                readPlainStringPart(reader, richString);
        }
    }

    int idx = m_stringList.size();
    m_stringTable[richString] = XlsxSharedStringInfo(idx, 0);
    m_stringList.append(richString);
}

void SharedStrings::readRichStringPart(XmlStreamReader &reader, RichString &richString)
{
    Q_ASSERT(reader.name() == QLatin1String("r"));

    QString text;
    Format *format=0;
    while (!(reader.name() == QLatin1String("r") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("rPr")) {
                //:Todo
            } else if (reader.name() == QLatin1String("t")) {
                text = reader.readElementText();
            }
        }
    }
    richString.addFragment(text, format);
}

void SharedStrings::readPlainStringPart(XmlStreamReader &reader, RichString &richString)
{
    Q_ASSERT(reader.name() == QLatin1String("t"));

    //QXmlStreamAttributes attributes = reader.attributes();

    QString text = reader.readElementText();
    richString.addFragment(text, 0);
}

bool SharedStrings::loadFromXmlFile(QIODevice *device)
{
    XmlStreamReader reader(device);
    int count = 0;
    while(!reader.atEnd()) {
         QXmlStreamReader::TokenType token = reader.readNext();
         if (token == QXmlStreamReader::StartElement) {
             if (reader.name() == QLatin1String("sst")) {
                 QXmlStreamAttributes attributes = reader.attributes();
                 count = attributes.value(QLatin1String("uniqueCount")).toString().toInt();
             } else if (reader.name() == QLatin1String("si")) {
                 readString(reader);
             }
         }
    }

    if (m_stringTable.size() != count) {
        qDebug("Error: Shared string count");
        return false;
    }

    return true;
}

bool SharedStrings::loadFromXmlData(const QByteArray &data)
{
    QBuffer buffer;
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);

    return loadFromXmlFile(&buffer);
}

} //namespace
