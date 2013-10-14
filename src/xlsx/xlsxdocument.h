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

#ifndef QXLSX_XLSXDOCUMENT_H
#define QXLSX_XLSXDOCUMENT_H

#include "xlsxglobal.h"
#include <QObject>
class QIODevice;
class QImage;

namespace QXlsx {

class Workbook;
class Worksheet;
class Package;
class Format;

class DocumentPrivate;
class Q_XLSX_EXPORT Document : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Document)

public:
    explicit Document(QObject *parent = 0);
    Document(const QString &xlsxName, QObject *parent=0);
    Document(QIODevice *device, QObject *parent=0);
    ~Document();

    Format *createFormat();
    int write(const QString cell, const QVariant &value, Format *format=0);
    int write(int row, int col, const QVariant &value, Format *format=0);
    int insertImage(int row, int column, const QImage &image, double xOffset=0, double yOffset=0, double xScale=1, double yScale=1);
    int mergeCells(const QString &range);
    int unmergeCells(const QString &range);
    bool setRow(int row, double height, Format* format=0, bool hidden=false);
    bool setColumn(int colFirst, int colLast, double width, Format* format=0, bool hidden=false);

    QString documentProperty(const QString &name) const;
    void setDocumentProperty(const QString &name, const QString &property);
    QStringList documentPropertyNames() const;

    Workbook *workbook() const;
    bool addWorksheet(const QString &name = QString());
    bool insertWorkSheet(int index, const QString &name = QString());
    Worksheet *activedWorksheet() const;
    int activedWorksheetIndex() const;
    void setActivedWorksheetIndex(int index);

    bool save();
    bool saveAs(const QString &xlsXname);
    bool saveAs(QIODevice *device);

private:
    friend class Package;
    Q_DISABLE_COPY(Document)
    DocumentPrivate * const d_ptr;
};

} // namespace QXlsx

#endif // QXLSX_XLSXDOCUMENT_H
