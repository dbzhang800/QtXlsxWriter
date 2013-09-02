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
#ifndef XLSXWORKSHEET_P_H
#define XLSXWORKSHEET_P_H
#include "xlsxworksheet.h"

namespace QXlsx {

struct XlsxCellData
{
    enum CellDataType {
        Blank,
        String,
        Number,
        Formula,
        ArrayFormula,
        Boolean,
        DateTime
    };
    XlsxCellData(const QVariant &data=QVariant(), CellDataType type=Blank, Format *format=0) :
        value(data), dataType(type), format(format)
    {

    }

    QVariant value;
    QString formula;
    CellDataType dataType;
    Format *format;
};

struct XlsxUrlData
{
    XlsxUrlData(int linkType=1, const QString &url=QString(), const QString &location=QString(), const QString &tip=QString()) :
        linkType(linkType), url(url), location(location), tip(tip)
    {

    }

    int linkType;
    QString url;
    QString location; //location string
    QString tip;
};

struct XlsxRowInfo
{
    XlsxRowInfo(double height, Format *format, bool hidden) :
        height(height), format(format), hidden(hidden)
    {

    }

    double height;
    Format *format;
    bool hidden;
};

struct XlsxColumnInfo
{
    XlsxColumnInfo(int column_min, int column_max, double width, Format *format, bool hidden) :
        column_min(column_min), column_max(column_max), width(width), format(format), hidden(hidden)
    {

    }
    int column_min;
    int column_max;
    double width;
    Format *format;
    bool hidden;
};

class WorksheetPrivate
{
    Q_DECLARE_PUBLIC(Worksheet)
public:
    WorksheetPrivate(Worksheet *p);
    ~WorksheetPrivate();
    int checkDimensions(int row, int col, bool ignore_row=false, bool ignore_col=false);
    QString generateDimensionString();
    void calculateSpans();
    void writeSheetData(XmlStreamWriter &writer);
    void writeCellData(XmlStreamWriter &writer, int row, int col, XlsxCellData *cell);
    void writeHyperlinks(XmlStreamWriter &writer);

    Workbook *workbook;
    QMap<int, QMap<int, XlsxCellData *> > cellTable;
    QMap<int, QMap<int, QString> > comments;
    QMap<int, QMap<int, XlsxUrlData *> > urlTable;
    QStringList externUrlList;
    QMap<int, XlsxRowInfo *> rowsInfo;
    QList<XlsxColumnInfo *> colsInfo;
    QMap<int, XlsxColumnInfo *> colsInfoHelper;//Not owns the XlsxColumnInfo

    int xls_rowmax;
    int xls_colmax;
    int xls_strmax;
    int dim_rowmin;
    int dim_rowmax;
    int dim_colmin;
    int dim_colmax;
    int previous_row;

    QMap<int, QString> row_spans;

    int outline_row_level;
    int outline_col_level;

    int default_row_height;
    bool default_row_zeroed;

    QString name;
    bool hidden;
    bool selected;
    bool right_to_left;
    bool show_zeros;

    Worksheet *q_ptr;
};

}
#endif // XLSXWORKSHEET_P_H
