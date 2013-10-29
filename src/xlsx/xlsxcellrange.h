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
#ifndef QXLSX_XLSXCELLRANGE_H
#define QXLSX_XLSXCELLRANGE_H
#include "xlsxglobal.h"

QT_BEGIN_NAMESPACE_XLSX

class Q_XLSX_EXPORT CellRange
{
public:
    CellRange();
    CellRange(int firstRow, int firstColumn, int lastRow, int lastColumn);
    CellRange(const QString &range);
    CellRange(const CellRange &other);
    ~CellRange();

    QString toString() const;
    inline int firstRow() const { return top; }
    inline int lastRow() const { return bottom; }
    inline int firstColumn() const { return left; }
    inline int lastColumn() const { return right; }
    inline int rowCount() const { return bottom - top + 1; }
    inline int columnCount() const { return right - left + 1; }

private:
    int top, left, bottom, right;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXCELLRANGE_H
