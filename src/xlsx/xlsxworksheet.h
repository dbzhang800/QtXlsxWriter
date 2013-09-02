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
#ifndef XLSXWORKSHEET_H
#define XLSXWORKSHEET_H

#include "xlsxglobal.h"
#include <QObject>
#include <QStringList>
#include <QMap>
#include <QVariant>
class QIODevice;
class QDateTime;
class QUrl;

namespace QXlsx {
class Package;
class Workbook;
class XmlStreamWriter;
class Format;

class WorksheetPrivate;
class Q_XLSX_EXPORT Worksheet : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Worksheet)
public:
    int write(const QString row_column, const QVariant &value, Format *format=0);
    int write(int row, int column, const QVariant &value, Format *format=0);
    int writeString(int row, int column, const QString &value, Format *format=0);
    int writeNumber(int row, int column, double value, Format *format=0);
    int writeFormula(int row, int column, const QString &formula, Format *format=0, double result=0);
    int writeBlank(int row, int column, Format *format=0);
    int writeBool(int row, int column, bool value, Format *format=0);
    int writeDateTime(int row, int column, const QDateTime& dt, Format *format=0);
    int writeUrl(int row, int column, const QUrl &url, Format *format=0, const QString &display=QString(), const QString &tip=QString());

    bool setRow(int row, double height, Format* format=0, bool hidden=false);
    bool setColumn(int colFirst, int colLast, double width, Format* format=0, bool hidden=false);

    void setRightToLeft(bool enable);
    void setZeroValuesHidden(bool enable);

private:
    friend class Package;
    friend class Workbook;
    Worksheet(const QString &sheetName, Workbook *parent=0);
    ~Worksheet();

    virtual bool isChartsheet() const;
    QString name() const;
    bool isHidden() const;
    bool isSelected() const;
    void setHidden(bool hidden);
    void setSelected(bool select);
    void saveToXmlFile(QIODevice *device);
    QStringList externUrlList() const;

    WorksheetPrivate * const d_ptr;
};

} //QXlsx

#endif // XLSXWORKSHEET_H
