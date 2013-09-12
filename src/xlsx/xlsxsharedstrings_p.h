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
#ifndef XLSXSHAREDSTRINGS_H
#define XLSXSHAREDSTRINGS_H

#include "xlsxglobal.h"
#include <QHash>
#include <QStringList>
#include <QSharedPointer>

class QIODevice;

namespace QXlsx {

class XlsxSharedStringInfo
{
public:
    XlsxSharedStringInfo(int index=0, int count = 1) :
        index(index), count(count)
    {
    }

    int index;
    int count;
};

class XLSX_AUTOTEST_EXPORT SharedStrings
{
public:
    SharedStrings();
    int count() const;
    
    int addSharedString(const QString &string);
    void removeSharedString(const QString &string);

    int getSharedStringIndex(const QString &string) const;
    QString getSharedString(int index) const;
    QStringList getSharedStrings() const;

    void saveToXmlFile(QIODevice *device) const;
    QByteArray saveToXmlData() const;
    static QSharedPointer<SharedStrings> loadFromXmlFile(QIODevice *device);
    static QSharedPointer<SharedStrings> loadFromXmlData(const QByteArray &data);

private:
    QHash<QString, XlsxSharedStringInfo> m_stringTable; //for fast lookup
    QStringList m_stringList;
    int m_stringCount;
};

}
#endif // XLSXSHAREDSTRINGS_H
