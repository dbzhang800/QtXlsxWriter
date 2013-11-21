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
#include <QList>
#include <QImage>
#include <QSharedPointer>

class QIODevice;

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;
class SharedStrings;
class Styles;
class Package;
class Drawing;
class Document;
class Theme;
class DocumentPrivate;

class WorkbookPrivate;
class Q_XLSX_EXPORT Workbook
{
    Q_DECLARE_PRIVATE(Workbook)
public:
    ~Workbook();

    QList<QSharedPointer<Worksheet> > worksheets() const;
    Worksheet *addWorksheet(const QString &name = QString());
    Worksheet *insertWorkSheet(int index, const QString &name = QString());
    int activeWorksheet() const;
    void setActiveWorksheet(int index);

//    void addChart();
    bool defineName(const QString &name, const QString &formula, const QString &comment=QString(), const QString &scope=QString());
    bool isDate1904() const;
    void setDate1904(bool date1904);
    bool isStringsToNumbersEnabled() const;
    void setStringsToNumbersEnabled(bool enable=true);
    QString defaultDateFormat() const;
    void setDefaultDateFormat(const QString &format);

private:
    friend class Package;
    friend class Worksheet;
    friend class WorksheetPrivate;
    friend class Document;
    friend class DocumentPrivate;

    Workbook();

    void saveToXmlFile(QIODevice *device);
    QByteArray saveToXmlData();
    bool loadFromXmlFile(QIODevice *device);
    bool loadFromXmlData(const QByteArray &data);

    SharedStrings *sharedStrings() const;
    Styles *styles();
    Theme *theme();
    QList<QImage> images();
    QList<Drawing *> drawings();
    void prepareDrawings();
    Worksheet *addWorksheet(const QString &name, int sheetId);
    WorkbookPrivate * const d_ptr;
};

QT_END_NAMESPACE_XLSX

#endif // XLSXWORKBOOK_H
