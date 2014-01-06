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
#include "xlsxutility_p.h"

#include <QString>
#include <QPoint>
#include <QRegularExpression>
#include <QMap>
#include <QStringList>
#include <QColor>
#include <QDateTime>
#include <QDebug>

namespace QXlsx {

int intPow(int x, int p)
{
  if (p == 0) return 1;
  if (p == 1) return x;

  int tmp = intPow(x, p/2);
  if (p%2 == 0) return tmp * tmp;
  else return x * tmp * tmp;
}

QStringList splitPath(const QString &path)
{
    int idx = path.lastIndexOf(QLatin1Char('/'));
    if (idx == -1)
        return QStringList()<<QStringLiteral(".")<<path;

    return QStringList()<<path.left(idx)<<path.mid(idx+1);
}

/*
 * Return the .rel file path based on filePath
 */
QString getRelFilePath(const QString &filePath)
{
    int idx = filePath.lastIndexOf(QLatin1Char('/'));
    if (idx == -1)
        return QString();

    return QString(filePath.left(idx) + QLatin1String("/_rels/")
                   + filePath.mid(idx+1) + QLatin1String(".rels"));
}

double datetimeToNumber(const QDateTime &dt, bool is1904)
{
    //Note, for number 0, Excel2007 shown as 1900-1-0, which should be 1899-12-31
    QDateTime epoch(is1904 ? QDate(1904, 1, 1): QDate(1899, 12, 31), QTime(0,0));

    double excel_time = epoch.msecsTo(dt) / (1000*60*60*24.0);
    if (!is1904 && excel_time > 59) {//31+28
        //Account for Excel erroneously treating 1900 as a leap year.
        excel_time += 1;
    }
    return excel_time;
}

double timeToNumber(const QTime &time)
{
    return QTime(0,0).msecsTo(time) / (1000*60*60*24.0);
}

QDateTime datetimeFromNumber(double num, bool is1904)
{
    if (!is1904 && num > 60)
        num = num - 1;

    qint64 msecs = static_cast<qint64>(num * 1000*60*60*24.0 + 0.5);
    QDateTime epoch(is1904 ? QDate(1904, 1, 1): QDate(1899, 12, 31), QTime(0,0));

    return epoch.addMSecs(msecs);
}

QPoint xl_cell_to_rowcol(const QString &cell_str)
{
    if (cell_str.isEmpty())
        return QPoint(-1, -1);
    QRegularExpression re(QStringLiteral("^([A-Z]{1,3})(\\d+)$"));
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

        int row = row_str.toInt();
        return QPoint(row, col);
    } else {
        return QPoint(-1, -1); //...
    }
}

QString xl_col_to_name(int col_num)
{
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

int xl_col_name_to_value(const QString &col_str)
{
    QRegularExpression re(QStringLiteral("^([A-Z]{1,3})$"));
    QRegularExpressionMatch match = re.match(col_str);
    if (match.hasMatch()) {
        int col = 0;
        int expn = 0;
        for (int i=col_str.size()-1; i>-1; --i) {
            col += (col_str[i].unicode() - 'A' + 1) * intPow(26, expn);
            expn++;
        }

        return col;
    }
    return -1;
}

QString xl_rowcol_to_cell(int row, int col, bool row_abs, bool col_abs)
{
    QString cell_str;
    if (col_abs)
        cell_str.append(QLatin1Char('$'));
    cell_str.append(xl_col_to_name(col));
    if (row_abs)
        cell_str.append(QLatin1Char('$'));
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
    return col_str + QString::number(row);
}

} //namespace QXlsx
