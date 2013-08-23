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
#include "xlsxutility_p.h"

#include <QString>
#include <QPoint>
#include <QRegularExpression>
#include <QMap>

namespace QXlsx {

int intPow(int x, int p)
{
  if (p == 0) return 1;
  if (p == 1) return x;

  int tmp = intPow(x, p/2);
  if (p%2 == 0) return tmp * tmp;
  else return x * tmp * tmp;
}

QPoint xl_cell_to_rowcol(const QString &cell_str)
{
    if (cell_str.isEmpty())
        return QPoint(0, 0);
    QRegularExpression re("^([A-Z]{1,3})(\\d+)$");
    QRegularExpressionMatch match = re.match(cell_str);
    if (match.hasMatch()) {
        QString col_str = match.captured(1);
        QString row_str = match.captured(2);
        int col = 0;
        int expn = 0;
        for (int i=col_str.size()-1; i>-1; --i) {
            col += (col_str[i].unicode() - 'A' + 1) * intPow(26, expn);
            expn++;
        }

        col--;
        int row = row_str.toInt() - 1;
        return QPoint(row, col);
    } else {
        return QPoint(-1, -1); //...
    }
}

QString xl_col_to_name(int col_num)
{
    col_num += 1; //Change to 1-index
    QString col_str;

    int remainder;
    while (col_num) {
        remainder = col_num % 26;
        if (remainder == 0)
            remainder = 26;
        col_str.prepend(QChar('A'+remainder-1));
        col_num = (col_num - 1) / 26;
    }

    return col_str;
}

QString xl_rowcol_to_cell(int row, int col, bool row_abs, bool col_abs)
{
    row += 1; //Change to 1-index
    QString cell_str;
    if (col_abs)
        cell_str.append("$");
    cell_str.append(xl_col_to_name(col));
    if (row_abs)
        cell_str.append("$");
    cell_str.append(QString::number(row));
    return cell_str;
}

QString xl_rowcol_to_cell_fast(int row, int col)
{
    static QMap<int, QString> col_cache;
    QString  col_str;
    if (col_cache.contains(col)) {
        col_str = col_cache[col];
    } else {
        col_str = xl_col_to_name(col);
        col_cache[col] = col_str;
    }
    return col_str + QString::number(row+1);
}

} //namespace QXlsx
