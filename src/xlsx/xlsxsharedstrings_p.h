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

#include <QObject>
#include <QHash>
#include <QStringList>

class QIODevice;

namespace QXlsx {

class SharedStrings : public QObject
{
    Q_OBJECT
public:
    explicit SharedStrings(QObject *parent = 0);
    int count() const;
    
public slots:
    int addSharedString(const QString &string);
    int getSharedStringIndex(const QString &string) const;
    QString getSharedString(int index) const;
    QStringList getSharedStrings() const;

    void saveToXmlFile(QIODevice *device) const;

private:
    QHash<QString, int> m_stringTable; //for fast lookup
    QStringList m_stringList;
    int m_stringCount;
};

}
#endif // XLSXSHAREDSTRINGS_H
