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
#include "xlsxrelationships_p.h"
#include "xlsxxmlwriter_p.h"
#include <QDir>
#include <QFile>

namespace QXlsx {

Relationships::Relationships(QObject *parent) :
    QObject(parent)
{
}

void Relationships::addDocumentRelationship(const QString &relativeType, const QString &target)
{
    QString type = QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + relativeType;
    addRelationship(type, target);
}

void Relationships::addMsPackageRelationship(const QString &relativeType, const QString &target)
{
    QString type = QStringLiteral("http://schemas.microsoft.com/office/2006/relationships") + relativeType;
    addRelationship(type, target);
}

void Relationships::addPackageRelationship(const QString &relativeType, const QString &target)
{
    QString type = QStringLiteral("http://schemas.openxmlformats.org/package/2006/relationships") + relativeType;
    addRelationship(type, target);
}

void Relationships::addWorksheetRelationship(const QString &relativeType, const QString &target, const QString &targetMode)
{
    QString type = QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + relativeType;
    addRelationship(type, target, targetMode);
}

void Relationships::addRelationship(const QString &type, const QString &target, const QString &targetMode)
{
    XlsxRelationship relation;
    relation.id = QStringLiteral("rId%1").arg(m_relationships.size()+1);
    relation.type = type;
    relation.target = target;
    relation.targetMode = targetMode;

    m_relationships.append(relation);
}

void Relationships::saveToXmlFile(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("Relationships"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/package/2006/relationships"));
    foreach (XlsxRelationship relation, m_relationships) {
        writer.writeStartElement(QStringLiteral("Relationship"));
        writer.writeAttribute(QStringLiteral("Id"), relation.id);
        writer.writeAttribute(QStringLiteral("Type"), relation.type);
        writer.writeAttribute(QStringLiteral("Target"), relation.target);
        if (!relation.targetMode.isNull())
            writer.writeAttribute(QStringLiteral("TargetMode"), relation.targetMode);
        writer.writeEndElement();
    }
    writer.writeEndElement();//Relationships
    writer.writeEndDocument();
}

} //namespace
