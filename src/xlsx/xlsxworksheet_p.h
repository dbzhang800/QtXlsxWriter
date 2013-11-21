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
#include "xlsxcell.h"
#include "xlsxdatavalidation.h"

#include <QImage>
#include <QSharedPointer>

namespace QXlsx {

class XmlStreamWriter;
class XmlStreamReader;
class SharedStrings;

struct XlsxUrlData
{
    enum LinkType
    {
        External,
        Internal
    };

    XlsxUrlData(LinkType linkType=External, const QString &url=QString(), const QString &location=QString(), const QString &tip=QString()) :
        linkType(linkType), url(url), location(location), tip(tip)
    {

    }

    LinkType linkType;
    QString url;
    QString location; //location string
    QString display;
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
    XlsxRowInfo(double height=0, const Format &format=Format(), bool hidden=false) :
        height(height), format(format), hidden(hidden), outlineLevel(0)
      , collapsed(false)
    {

    }

    double height;
    Format format;
    bool hidden;
    int outlineLevel;
    bool collapsed;
};

struct XlsxColumnInfo
{
    XlsxColumnInfo(int firstColumn=0, int lastColumn=1, double width=0, const Format &format=Format(), bool hidden=false) :
        firstColumn(firstColumn), lastColumn(lastColumn), width(width), format(format), hidden(hidden)
      , outlineLevel(0), collapsed(false)
    {

    }
    int firstColumn;
    int lastColumn;
    double width;
    Format format;
    bool hidden;
    int outlineLevel;
    bool collapsed;
};

class XLSX_AUTOTEST_EXPORT WorksheetPrivate
{
    Q_DECLARE_PUBLIC(Worksheet)
public:
    WorksheetPrivate(Worksheet *p);
    ~WorksheetPrivate();
    int checkDimensions(int row, int col, bool ignore_row=false, bool ignore_col=false);
    Format cellFormat(int row, int col) const;
    QString generateDimensionString();
    void calculateSpans();
    void splitColsInfo(int colFirst, int colLast);

    void writeSheetData(XmlStreamWriter &writer);
    void writeCellData(XmlStreamWriter &writer, int row, int col, QSharedPointer<Cell> cell);
    void writeMergeCells(XmlStreamWriter &writer);
    void writeHyperlinks(XmlStreamWriter &writer);
    void writeDrawings(XmlStreamWriter &writer);
    void writeDataValidation(XmlStreamWriter &writer);
    int rowPixelsSize(int row);
    int colPixelsSize(int col);
    XlsxObjectPositionData objectPixelsPosition(int col_start, int row_start, double x1, double y1, double width, double height);
    XlsxObjectPositionData pixelsToEMUs(const XlsxObjectPositionData &data);

    QSharedPointer<Cell> readNumericCellData(XmlStreamReader &reader);
    void readSheetData(XmlStreamReader &reader);
    void readColumnsInfo(XmlStreamReader &reader);
    void readMergeCells(XmlStreamReader &reader);
    void readDataValidations(XmlStreamReader &reader);
    void readDataValidation(XmlStreamReader &reader);
    void readSheetViews(XmlStreamReader &reader);

    SharedStrings *sharedStrings() const;

    Worksheet *q_ptr;
    Workbook *workbook;
    Drawing *drawing;
    QMap<int, QMap<int, QSharedPointer<Cell> > > cellTable;
    QMap<int, QMap<int, QString> > comments;
    QMap<int, QMap<int, XlsxUrlData *> > urlTable;
    QList<CellRange> merges;
    QStringList externUrlList;
    QStringList externDrawingList;
    QList<XlsxImageData *> imageList;
    QMap<int, QSharedPointer<XlsxRowInfo> > rowsInfo;
    QMap<int, QSharedPointer<XlsxColumnInfo> > colsInfo;
    QMap<int, QSharedPointer<XlsxColumnInfo> > colsInfoHelper;
    QList<QPair<QString, QString> > drawingLinks;

    QList<DataValidation> dataValidationsList;

    int xls_rowmax;
    int xls_colmax;
    int xls_strmax;
    CellRange dimension;
    int previous_row;

    QMap<int, QString> row_spans;
    QMap<int, double> row_sizes;
    QMap<int, double> col_sizes;

    int outline_row_level;
    int outline_col_level;

    int default_row_height;
    bool default_row_zeroed;

    QString name;
    int id;
    bool hidden;

    bool windowProtection;
    bool showFormulas;
    bool showGridLines;
    bool showRowColHeaders;
    bool showZeros;
    bool rightToLeft;
    bool tabSelected;
    bool showRuler;
    bool showOutlineSymbols;
    bool showWhiteSpace;
};

}
#endif // XLSXWORKSHEET_P_H
