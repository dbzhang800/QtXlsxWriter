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

bool parseXsdBoolean(const QString &value, bool defaultValue)
{
    if (value == QLatin1String("1") || value == QLatin1String("true"))
        return true;
    if (value == QLatin1String("0") || value == QLatin1String("false"))
        return false;
    return defaultValue;
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

#if QT_VERSION >= 0x050200
    if (dt.isDaylightTime())    // Add one hour if the date is Daylight
        excel_time += 1.0 / 24.0;
#endif

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

    QDateTime dt = epoch.addMSecs(msecs);

#if QT_VERSION >= 0x050200
    // Remove one hour to see whether the date is Daylight
    QDateTime dt2 = dt.addMSecs(-3600);
    if (dt2.isDaylightTime())
        return dt2;
#endif

    return dt;
}

/*
  Creates a valid sheet name
    minimum length is 1
    maximum length is 31
    doesn't contain special chars: / \ ? * ] [ :
    Sheet names must not begin or end with ' (apostrophe)

  Invalid characters are replaced by one space character ' '.
 */
QString createSafeSheetName(const QString &nameProposal)
{
    if (nameProposal.isEmpty())
        return QString();

    QString ret = nameProposal;
    if (nameProposal.contains(QRegularExpression(QStringLiteral("[/\\\\?*\\][:]+"))))
        ret.replace(QRegularExpression(QStringLiteral("[/\\\\?*\\][:]+")), QStringLiteral(" "));
    while(ret.contains(QRegularExpression(QStringLiteral("^\\s*'\\s*|\\s*'\\s*$"))))
        ret.remove(QRegularExpression(QStringLiteral("^\\s*'\\s*|\\s*'\\s*$")));
    ret = ret.trimmed();
    if (ret.size() > 31)
        ret = ret.left(31);
    return ret;
}

/*
 * whether the string s starts or ends with space
 */
bool isSpaceReserveNeeded(const QString &s)
{
    QString spaces(QStringLiteral(" \t\n\r"));
    return !s.isEmpty() && (spaces.contains(s.at(0))||spaces.contains(s.at(s.length()-1)));
}

} //namespace QXlsx
