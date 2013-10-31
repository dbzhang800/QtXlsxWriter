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
#ifndef XLSXUTILITY_H
#define XLSXUTILITY_H

#include "xlsxglobal.h"
class QPoint;
class QString;
class QStringList;
class QColor;
class QDateTime;

namespace QXlsx {

XLSX_AUTOTEST_EXPORT int intPow(int x, int p);
XLSX_AUTOTEST_EXPORT QStringList splitPath(const QString &path);
XLSX_AUTOTEST_EXPORT QColor fromARGBString(const QString &c);
XLSX_AUTOTEST_EXPORT double datetimeToNumber(const QDateTime &dt, bool is1904=false);
XLSX_AUTOTEST_EXPORT QDateTime datetimeFromNumber(double num, bool is1904=false);

XLSX_AUTOTEST_EXPORT QPoint xl_cell_to_rowcol(const QString &cell_str);
XLSX_AUTOTEST_EXPORT QString xl_col_to_name(int col_num);
XLSX_AUTOTEST_EXPORT int xl_col_name_to_value(const QString &col_str);
XLSX_AUTOTEST_EXPORT QString xl_rowcol_to_cell(int row, int col, bool row_abs=false, bool col_abs=false);
XLSX_AUTOTEST_EXPORT QString xl_rowcol_to_cell_fast(int row, int col);

} //QXlsx
#endif // XLSXUTILITY_H
