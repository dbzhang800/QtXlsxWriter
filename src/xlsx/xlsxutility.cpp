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
#include "xlsxcellreference.h"

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

/*
 * Convert shared formula for non-root cells.
 *
 * For example, if "B1:B10" have shared formula "=A1*A1", this function will return "=A2*A2"
 * for "B2" cell, "=A3*A3" for "B3" cell, etc.
 *
 * Note, the formula "=A1*A1" for B1 can also be written as "=RC[-1]*RC[-1]", which is the same
 * for all other cells. In other words, this formula is shared.
 *
 * For long run, we need a formula parser.
 */
QString convertSharedFormula(const QString &rootFormula, const CellReference &rootCell, const CellReference &cell)
{
    //Find all the "[A-Z]+[0-9]+" patterns in the rootFormula.
    QList<QPair<QString, bool> > segments;

    QString segment;
    bool inQuote = false;
    int cellFlag = 0; //-1, 0, 1, 2 ==> Invalid, Empty, A-Z ready, A1 ready
    foreach (QChar ch, rootFormula) {
        if (inQuote) {
            segment.append(ch);
            if (ch == QLatin1Char('"')) {
                segments.append(qMakePair(segment, false));
                segment = QString();
                inQuote = false;
                cellFlag = 0;
            }
        } else {
            if (ch == QLatin1Char('"')) {
                segments.append(qMakePair(segment, false));
                segment = QString(ch);
                inQuote = true;
            } else if (ch >= QLatin1Char('A') && ch <=QLatin1Char('Z')) {
                if (cellFlag == 0 || cellFlag == 1) {
                    segment.append(ch);
                } else {
                    segments.append(qMakePair(segment, (cellFlag == 2)));
                    segment = QString(ch); //start new "A1" segment
                }
                cellFlag = 1;
            } else if (ch >= QLatin1Char('0') && ch <=QLatin1Char('9')) {
                segment.append(ch);
                if (cellFlag == 1)
                    cellFlag = 2;
            } else {
                if (cellFlag == 2) {
                    segments.append(qMakePair(segment, true)); //find one "A1" segment
                    segment = QString(ch);
                } else {
                    segment.append(ch);
                }
                cellFlag = -1;
            }
        }
    }

    if (!segment.isEmpty())
        segments.append(qMakePair(segment, (cellFlag == 2)));

    //Replace "A1" segment with proper one.
    QStringList result;
    typedef QPair<QString, bool> PairType;
    foreach (PairType p, segments) {
        if (p.second) {
            CellReference oldRef(p.first);
            CellReference newRef(oldRef.row()-rootCell.row()+cell.row(),
                                 oldRef.column()-rootCell.column()+cell.column());
            result.append(newRef.toString());
        } else {
            result.append(p.first);
        }
    }

    //OK
    return result.join(QString());
}

} //namespace QXlsx
