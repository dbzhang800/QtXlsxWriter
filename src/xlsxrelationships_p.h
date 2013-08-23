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
#ifndef XLSXRELATIONSHIPS_H
#define XLSXRELATIONSHIPS_H

#include <QObject>
#include <QList>
class QIODevice;

namespace QXlsx {

struct XlsxRelationship
{
    QString id;
    QString type;
    QString target;
    QString targetMode;
};

class Relationships : public QObject
{
    Q_OBJECT
public:
    explicit Relationships(QObject *parent = 0);
    
signals:
    
public slots:
    void addDocumentRelationship(const QString &relativeType, const QString &target);
    void addPackageRelationship(const QString &relativeType, const QString &target);
    void addMsPackageRelationship(const QString &relativeType, const QString &target);
    void addWorksheetRelationship(const QString &relativeType, const QString &target, const QString &targetMode=QString());

    void saveToXmlFile(QIODevice *device);
private:
    void addRelationship(const QString &type, const QString &target, const QString &targetMode=QString());
    QList<XlsxRelationship> m_relationships;
};

}
#endif // XLSXRELATIONSHIPS_H
