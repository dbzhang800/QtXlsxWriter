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
#ifndef XLSXWORKBOOK_P_H
#define XLSXWORKBOOK_P_H

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

#include "xlsxworkbook.h"
#include "xlsxtheme_p.h"
#include "xlsxrelationships_p.h"

#include <QSharedPointer>
#include <QPair>
#include <QStringList>

namespace QXlsx {

struct XlsxSheetItemInfo
{
    XlsxSheetItemInfo(){}

    QString name;
    int sheetId;
    QString rId;
    QString state;
};

struct XlsxDefineNameData
{
    XlsxDefineNameData()
        :sheetId(-1)
    {}
    XlsxDefineNameData(const QString &name, const QString &formula, const QString &comment, int sheetId=-1)
        :name(name), formula(formula), comment(comment), sheetId(sheetId)
    {

    }
    QString name;
    QString formula;
    QString comment;
    //using internal sheetId, instead of the localSheetId(order in the workbook)
    int sheetId;
};

class WorkbookPrivate
{
    Q_DECLARE_PUBLIC(Workbook)
public:
    WorkbookPrivate(Workbook *q);

    Workbook *q_ptr;
    mutable Relationships relationships;

    QSharedPointer<SharedStrings> sharedStrings;
    QList<QSharedPointer<Worksheet> > worksheets;
    QStringList worksheetNames;
    QSharedPointer<Styles> styles;
    QSharedPointer<Theme> theme;
    QList<QImage> images;
    QList<Drawing *> drawings;
    QList<XlsxDefineNameData> definedNamesList;

    QList<XlsxSheetItemInfo> sheetItemInfoList;//Data from xml file

    bool strings_to_numbers_enabled;
    bool date1904;
    QString defaultDateFormat;

    int x_window;
    int y_window;
    int window_width;
    int window_height;

    int activesheetIndex;
    int firstsheet;
    int table_count;

    //Used to generate new sheet name and id
    int last_sheet_index;
    int last_sheet_id;
};

}

#endif // XLSXWORKBOOK_P_H
