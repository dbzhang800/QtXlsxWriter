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
#ifndef XLSXWORKBOOK_H
#define XLSXWORKBOOK_H

#include "xlsxglobal.h"
#include <QObject>
#include <QList>
#include <QImage>
class QIODevice;

namespace QXlsx {

class Worksheet;
class Format;
class SharedStrings;
class Styles;
class Package;
class Drawing;

class WorkbookPrivate;
class Q_XLSX_EXPORT Workbook : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Workbook)
public:
    Workbook(QObject *parent=0);
    ~Workbook();

    QList<Worksheet *> worksheets() const;
    Worksheet *addWorksheet(const QString &name = QString());
    Worksheet *insertWorkSheet(int index, const QString &name = QString());
    int activedWorksheet() const;
    void setActivedWorksheet(int index);

    Format *addFormat();
//    void addChart();
    void defineName(const QString &name, const QString &formula);
    bool isDate1904() const;
    void setDate1904(bool date1904);
    bool isStringsToNumbersEnabled() const;
    void setStringsToNumbersEnabled(bool enable=true);

#ifdef Q_QDOC
    bool setProperty(const char *key, const QString &value);
#endif

    bool save(const QString &name);

private:
    friend class Package;
    friend class Worksheet;

    SharedStrings *sharedStrings();
    Styles *styles();
    QList<QImage> images();
    QList<Drawing *> drawings();
    void prepareDrawings();
    void saveToXmlFile(QIODevice *device);

    WorkbookPrivate * const d_ptr;
};

} //QXlsx

#endif // XLSXWORKBOOK_H
