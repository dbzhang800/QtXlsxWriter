/****************************************************************************
** Copyright (c) 2013-2014 Debao Zhang <hello@debao.me>
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

#ifndef QXLSX_XLSXOOXMLFILE_H
#define QXLSX_XLSXOOXMLFILE_H

#include "xlsxglobal.h"

class QIODevice;
class QByteArray;

QT_BEGIN_NAMESPACE_XLSX

class OOXmlFilePrivate;

class Q_XLSX_EXPORT OOXmlFile
{
    Q_DECLARE_PRIVATE(OOXmlFile)
public:
    virtual ~OOXmlFile();

    virtual void saveToXmlFile(QIODevice *device) const = 0;
    virtual bool loadFromXmlFile(QIODevice *device) = 0;

    virtual QByteArray saveToXmlData() const;
    virtual bool loadFromXmlData(const QByteArray &data);

    void setFilePath(const QString path);
    QString filePath() const;
protected:
    OOXmlFile();
    OOXmlFile(OOXmlFilePrivate *d);

    OOXmlFilePrivate *d_ptr;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXOOXMLFILE_H
