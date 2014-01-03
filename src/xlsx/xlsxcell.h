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
#ifndef QXLSX_XLSXCELL_H
#define QXLSX_XLSXCELL_H

#include "xlsxglobal.h"
#include "xlsxformat.h"
#include <QVariant>

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;
class Format;
class CellPrivate;
class WorksheetPrivate;

class Q_XLSX_EXPORT Cell
{
    Q_DECLARE_PRIVATE(Cell)
public:
    enum DataType {
        Blank,
        String,
        Numeric,
        Formula,
        Boolean,
        Error,
        InlineString,
        ArrayFormula
    };

    DataType dataType() const;
    QVariant value() const;
    Format format() const;
    QString formula() const;

    bool isDateTime() const;
    QDateTime dateTime() const;

    bool isRichString() const;

    ~Cell();
private:
    friend class Worksheet;
    friend class WorksheetPrivate;

    Cell(const QVariant &data=QVariant(), DataType type=Blank, const Format &format=Format(), Worksheet *parent=0);
    Cell(const Cell * const cell);
    CellPrivate * const d_ptr;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXCELL_H
