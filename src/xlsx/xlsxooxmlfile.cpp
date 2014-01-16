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

#include "xlsxooxmlfile.h"
#include "xlsxooxmlfile_p.h"

#include <QBuffer>
#include <QByteArray>

QT_BEGIN_NAMESPACE_XLSX

OOXmlFilePrivate::OOXmlFilePrivate(OOXmlFile *q)
    :q_ptr(q)
{

}

/*!
 * \internal
 *
 * \class OOXmlFile
 *
 * Base class of all the ooxml part file.
 */

OOXmlFile::OOXmlFile()
    :d_ptr(new OOXmlFilePrivate(this))
{
}

OOXmlFile::OOXmlFile(OOXmlFilePrivate *d)
    :d_ptr(d)
{

}

OOXmlFile::~OOXmlFile()
{
    delete d_ptr;
}

QByteArray OOXmlFile::saveToXmlData() const
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);

    return data;
}

bool OOXmlFile::loadFromXmlData(const QByteArray &data)
{
    QBuffer buffer;
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);

    return loadFromXmlFile(&buffer);
}

/*!
 * \internal
 */
void OOXmlFile::setFilePath(const QString path)
{
    Q_D(OOXmlFile);
    d->filePathInPackage = path;
}

/*!
 * \internal
 */
QString OOXmlFile::filePath() const
{
    Q_D(const OOXmlFile);
    return d->filePathInPackage;
}

QT_END_NAMESPACE_XLSX
