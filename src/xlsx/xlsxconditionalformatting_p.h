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

#ifndef XLSXCONDITIONALFORMATTING_P_H
#define XLSXCONDITIONALFORMATTING_P_H

//
//  W A R N I N G
//  -------------
//
// This file is not part of the Qt Xlsx API.  It exists for the convenience
// of the Qt Xlsx.  This header file may change from
// version to version without notice, or even be removed.
//
// We mean it.
//

#include "xlsxConditionalFormatting.h"
#include "xlsxformat.h"
#include <QSharedData>
#include <QSharedPointer>
#include <QMap>

QT_BEGIN_NAMESPACE_XLSX

class XlsxCfRuleData
{
public:
    enum Attribute {
        A_type,
        A_dxfId,
        //A_priority,
        A_stopIfTrue,
        A_aboveAverage,
        A_percent,
        A_bottom,
        A_operator,
        A_text,
        A_timePeriod,
        A_rank,
        A_stdDev,
        A_equalAverage,

        A_dxfFormat,
        A_formula1,
        A_formula2,
        A_formula3,
        A_formula1_temp
    };

    XlsxCfRuleData()
        :priority(1)
    {}

    int priority;
    Format dxfFormat;
    QMap<int, QVariant> attrs;
};

class ConditionalFormattingPrivate : public QSharedData
{
public:
    ConditionalFormattingPrivate();
    ConditionalFormattingPrivate(const ConditionalFormattingPrivate &other);
    ~ConditionalFormattingPrivate();

    QList<QSharedPointer<XlsxCfRuleData> >cfRules;
    QList<CellRange> ranges;
};

QT_END_NAMESPACE_XLSX
#endif // XLSXCONDITIONALFORMATTING_P_H
