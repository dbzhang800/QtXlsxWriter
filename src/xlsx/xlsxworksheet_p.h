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
#include "xlsxglobal.h"
#include "xlsxworksheet.h"

#include <QImage>
#include <QSharedPointer>

namespace QXlsx {

class XmlStreamWriter;
class XmlStreamReader;

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

struct XlsxImageData
{
    XlsxImageData(int row, int col, const QImage &image, const QPointF &offset, double xScale, double yScale) :
        row(row), col(col), image(image), offset(offset), xScale(xScale), yScale(yScale)
    {
    }

    int row;
    int col;
    QImage image;
    QPointF offset;
    double xScale;
    double yScale;
};

struct XlsxCellRange
{
    int row_begin;
    int row_end;
    int column_begin;
    int column_end;

    bool operator ==(const XlsxCellRange &other) const {
        return row_begin==other.row_begin && row_end==other.row_end
                && column_begin == other.column_begin && column_end==other.column_end;
    }
    bool operator !=(const XlsxCellRange &other) const {
        return row_begin!=other.row_begin || row_end!=other.row_end
                || column_begin != other.column_begin || column_end!=other.column_end;
    }
};

/*
     The vertices that define the position of a graphical object
     within the worksheet in pixels.

             +------------+------------+
             |     A      |      B     |
       +-----+------------+------------+
       |     |(x1,y1)     |            |
       |  1  |(A1)._______|______      |
       |     |    |              |     |
       |     |    |              |     |
       +-----+----|    OBJECT    |-----+
       |     |    |              |     |
       |  2  |    |______________.     |
       |     |            |        (B2)|
       |     |            |     (x2,y2)|
       +---- +------------+------------+

     Example of an object that covers some of the area from cell A1 to  B2.

     Based on the width and height of the object we need to calculate 8 vars:

         col_start, row_start, col_end, row_end, x1, y1, x2, y2.

     We also calculate the absolute x and y position of the top left vertex of
     the object. This is required for images.

     The width and height of the cells that the object occupies can be
     variable and have to be taken into account.
*/
struct XlsxObjectPositionData
{
    int col_start;
    double x1;
    int row_start;
    double y1;
    int col_end;
    double x2;
    int row_end;
    double y2;
    double width;
    double height;
    double x_abs;
    double y_abs;
};

struct XlsxRowInfo
{
    XlsxRowInfo(double height=0, Format *format=0, bool hidden=false) :
        height(height), format(format), hidden(hidden)
    {

    }

    double height;
    Format *format;
    bool hidden;
};

struct XlsxColumnInfo
{
    XlsxColumnInfo(int column_min=0, int column_max=1, double width=0, Format *format=0, bool hidden=false) :
        column_min(column_min), column_max(column_max), width(width), format(format), hidden(hidden)
    {

    }
    int column_min;
    int column_max;
    double width;
    Format *format;
    bool hidden;
};

class XLSX_AUTOTEST_EXPORT WorksheetPrivate
{
    Q_DECLARE_PUBLIC(Worksheet)
public:
    WorksheetPrivate(Worksheet *p);
    ~WorksheetPrivate();
    int checkDimensions(int row, int col, bool ignore_row=false, bool ignore_col=false);
    QString generateDimensionString();
    void calculateSpans();
    void writeSheetData(XmlStreamWriter &writer);
    void writeCellData(XmlStreamWriter &writer, int row, int col, QSharedPointer<XlsxCellData> cell);
    void writeMergeCells(XmlStreamWriter &writer);
    void writeHyperlinks(XmlStreamWriter &writer);
    void writeDrawings(XmlStreamWriter &writer);
    int rowPixelsSize(int row);
    int colPixelsSize(int col);
    XlsxObjectPositionData objectPixelsPosition(int col_start, int row_start, double x1, double y1, double width, double height);
    XlsxObjectPositionData pixelsToEMUs(const XlsxObjectPositionData &data);

    void readSheetData(XmlStreamReader &reader);
    void readColumnsInfo(XmlStreamReader &reader);
    void readMergeCells(XmlStreamReader &reader);

    Workbook *workbook;
    Drawing *drawing;
    QMap<int, QMap<int, QSharedPointer<XlsxCellData> > > cellTable;
    QMap<int, QMap<int, QString> > comments;
    QMap<int, QMap<int, XlsxUrlData *> > urlTable;
    QList<XlsxCellRange> merges;
    QStringList externUrlList;
    QStringList externDrawingList;
    QList<XlsxImageData *> imageList;
    QMap<int, QSharedPointer<XlsxRowInfo> > rowsInfo;
    QList<QSharedPointer<XlsxColumnInfo> > colsInfo;
    QMap<int, QSharedPointer<XlsxColumnInfo> > colsInfoHelper;
    QList<QPair<QString, QString> > drawingLinks;

    int xls_rowmax;
    int xls_colmax;
    int xls_strmax;
    int dim_rowmin;
    int dim_rowmax;
    int dim_colmin;
    int dim_colmax;
    int previous_row;

    QMap<int, QString> row_spans;
    QMap<int, double> row_sizes;
    QMap<int, double> col_sizes;

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
