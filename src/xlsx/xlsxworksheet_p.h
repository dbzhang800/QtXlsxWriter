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
#ifndef XLSXWORKSHEET_P_H
#define XLSXWORKSHEET_P_H

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

#include "xlsxworksheet.h"
#include "xlsxabstractsheet_p.h"
#include "xlsxcell.h"
#include "xlsxdatavalidation.h"
#include "xlsxconditionalformatting.h"
#include "xlsxcellformula.h"

#include <QImage>
#include <QSharedPointer>
#include <QRegularExpression>

class QXmlStreamWriter;
class QXmlStreamReader;

namespace QXlsx {

const int XLSX_ROW_MAX = 1048576;
const int XLSX_COLUMN_MAX = 16384;
const int XLSX_STRING_MAX = 32767;

class SharedStrings;

struct XlsxHyperlinkData
{
    enum LinkType
    {
        External,
        Internal
    };

    XlsxHyperlinkData(LinkType linkType=External, const QString &target=QString(), const QString &location=QString()
            , const QString &display=QString(), const QString &tip=QString())
        :linkType(linkType), target(target), location(location), display(display), tooltip(tip)
    {

    }

    LinkType linkType;
    QString target; //For External link
    QString location;
    QString display;
    QString tooltip;
};

// ECMA-376 Part1 18.3.1.81
struct XlsxSheetFormatProps
{
    XlsxSheetFormatProps(int baseColWidth = 8,
                         bool customHeight = false,
                         double defaultColWidth = 0.0,
                         double defaultRowHeight = 15,
                         quint8 outlineLevelCol = 0,
                         quint8 outlineLevelRow = 0,
                         bool thickBottom = false,
                         bool thickTop = false,
                         bool zeroHeight = false) :
        baseColWidth(baseColWidth),
        customHeight(customHeight),
        defaultColWidth(defaultColWidth),
        defaultRowHeight(defaultRowHeight),
        outlineLevelCol(outlineLevelCol),
        outlineLevelRow(outlineLevelRow),
        thickBottom(thickBottom),
        thickTop(thickTop),
        zeroHeight(zeroHeight) {
    }

    int baseColWidth;
    bool customHeight;
    double defaultColWidth;
    double defaultRowHeight;
    quint8 outlineLevelCol;
    quint8 outlineLevelRow;
    bool thickBottom;
    bool thickTop;
    bool zeroHeight;
};

struct XlsxRowInfo
{
    XlsxRowInfo(double height=0, const Format &format=Format(), bool hidden=false) :
        customHeight(false), height(height), format(format), hidden(hidden), outlineLevel(0)
      , collapsed(false)
    {

    }

    bool customHeight;
    double height;
    Format format;
    bool hidden;
    int outlineLevel;
    bool collapsed;
};

struct XlsxColumnInfo
{
    XlsxColumnInfo(int firstColumn=0, int lastColumn=1, double width=0, const Format &format=Format(), bool hidden=false) :
        firstColumn(firstColumn), lastColumn(lastColumn), customWidth(false), width(width), format(format), hidden(hidden)
      , outlineLevel(0), collapsed(false)
    {

    }
    int firstColumn;
    int lastColumn;
    bool customWidth;
    double width;    
    Format format;
    bool hidden;
    int outlineLevel;
    bool collapsed;
};

struct XlsxPrintOptions
{
    XlsxPrintOptions(bool horizontalCentered=false, bool verticalCentered=false, bool headings=false, bool gridLines=false, bool gridLinesSet=true) :
        horizontalCentered(horizontalCentered), verticalCentered(verticalCentered), headings(headings), gridLines(gridLines), gridLinesSet(gridLinesSet)
    {

    }
    
    bool horizontalCentered;
    bool verticalCentered;
    bool headings;
    bool gridLines;
    bool gridLinesSet;
};

struct XlsxPageMargins
{
    // default values are from "Normal" setting of Microsoft Office 2010
    XlsxPageMargins(double left=0.7, double right=0.7, double top=0.75, double bottom=0.75, double header=0.3, double footer=0.3) :
        left(left), right(right), top(top), bottom(bottom), header(header), footer(footer)
    {

    }
    
    double left;
    double right;
    double top;
    double bottom;
    double header;
    double footer;
};

struct XlsxPageSetup
{
    
    static QString pageOrderString(Worksheet::PrintPageOrder pageOrder)
    {
        switch (pageOrder) {
            case Worksheet::DownThenOver: return "downThenOver";
            case Worksheet::OverThenDown: return "overThenDown";
            default:                      return "error";
        }
    }
    
    static QString orientationString(Worksheet::PrintOrientation orientation)
    {
        switch (orientation) {
            case Worksheet::Default:   return "default";
            case Worksheet::Portrait:  return "portrait";
            case Worksheet::Landscape: return "landscape";
            default:                   return "error";
        }
    }
    
    static QString cellCommentsString(Worksheet::PrintCellComments cellComments)
    {
        switch (cellComments) {
            case Worksheet::None:        return "none";
            case Worksheet::AsDisplayed: return "asDisplayed";
            case Worksheet::AtEnd:       return "atEnd";
            default:                     return "error";
        }
    }
    
    static QString errorsString(Worksheet::PrintErrors errors)
    {
        switch (errors) {
            case Worksheet::Displayed: return "displayed";
            case Worksheet::Blank:     return "blank";
            case Worksheet::Dash:      return "dash";
            case Worksheet::NA:        return "NA";
            default:                   return "error";
        }
    }
    
    // defaults are from the XMLSchema
    XlsxPageSetup() :
        paperSize(1), scale(100), firstPageNumber(1), fitToWidth(1), fitToHeight(1), pageOrder(Worksheet::DownThenOver),
        orientation(Worksheet::Default), usePrinterDefaults(true), blackAndWhite(false), draft(false), cellComments(Worksheet::None),
        useFirstPageNumber(false), errors(Worksheet::Displayed), horizontalDpi(600), verticalDpi(600), copies(1), rID()
    {

    }
    
    quint32 paperSize;
    quint32 scale;
    quint32 firstPageNumber;
    quint32 fitToWidth;
    quint32 fitToHeight;
    Worksheet::PrintPageOrder pageOrder;
    Worksheet::PrintOrientation orientation;
    bool usePrinterDefaults;
    bool blackAndWhite;
    bool draft;
    Worksheet::PrintCellComments cellComments;
    bool useFirstPageNumber;
    Worksheet::PrintErrors errors;
    quint32 horizontalDpi;
    quint32 verticalDpi;
    quint32 copies;
    QString rID;
};

class XLSX_AUTOTEST_EXPORT WorksheetPrivate : public AbstractSheetPrivate
{
    Q_DECLARE_PUBLIC(Worksheet)
public:
    WorksheetPrivate(Worksheet *p, Worksheet::CreateFlag flag);
    ~WorksheetPrivate();
    int checkDimensions(int row, int col, bool ignore_row=false, bool ignore_col=false);
    Format cellFormat(int row, int col) const;
    QString generateDimensionString() const;
    void calculateSpans() const;
    void splitColsInfo(int colFirst, int colLast);
    void validateDimension();

    void saveXmlSheetData(QXmlStreamWriter &writer) const;
    void saveXmlCellData(QXmlStreamWriter &writer, int row, int col, QSharedPointer<Cell> cell) const;
    void saveXmlMergeCells(QXmlStreamWriter &writer) const;
    void saveXmlHyperlinks(QXmlStreamWriter &writer) const;
    void saveXmlDrawings(QXmlStreamWriter &writer) const;
    void saveXmlDataValidations(QXmlStreamWriter &writer) const;
    void saveXmlPrintOptions(QXmlStreamWriter &writer) const;
    void saveXmlPageMargins(QXmlStreamWriter &writer) const;
    void saveXmlPageSetup(QXmlStreamWriter &writer) const;
    int rowPixelsSize(int row) const;
    int colPixelsSize(int col) const;

    void loadXmlSheetData(QXmlStreamReader &reader);
    void loadXmlColumnsInfo(QXmlStreamReader &reader);
    void loadXmlMergeCells(QXmlStreamReader &reader);
    void loadXmlDataValidations(QXmlStreamReader &reader);
    void loadXmlSheetFormatProps(QXmlStreamReader &reader);
    void loadXmlSheetViews(QXmlStreamReader &reader);
    void loadXmlHyperlinks(QXmlStreamReader &reader);
    void loadXmlPrintOptions(QXmlStreamReader &reader);
    void loadXmlPageMargins(QXmlStreamReader &reader);
    void loadXmlPageSetup(QXmlStreamReader &reader);

    QList<QSharedPointer<XlsxRowInfo> > getRowInfoList(int rowFirst, int rowLast);
    QList <QSharedPointer<XlsxColumnInfo> > getColumnInfoList(int colFirst, int colLast);
    QList<int> getColumnIndexes(int colFirst, int colLast);
    bool isColumnRangeValid(int colFirst, int colLast);

    SharedStrings *sharedStrings() const;

    QMap<int, QMap<int, QSharedPointer<Cell> > > cellTable;
    QMap<int, QMap<int, QString> > comments;
    QMap<int, QMap<int, QSharedPointer<XlsxHyperlinkData> > > urlTable;
    QList<CellRange> merges;
    QMap<int, QSharedPointer<XlsxRowInfo> > rowsInfo;
    QMap<int, QSharedPointer<XlsxColumnInfo> > colsInfo;
    QMap<int, QSharedPointer<XlsxColumnInfo> > colsInfoHelper;

    QList<DataValidation> dataValidationsList;
    QList<ConditionalFormatting> conditionalFormattingList;
    QMap<int, CellFormula> sharedFormulaMap;

    CellRange dimension;
    int previous_row;

    mutable QMap<int, QString> row_spans;
    QMap<int, double> row_sizes;
    QMap<int, double> col_sizes;

    int outline_row_level;
    int outline_col_level;

    int default_row_height;
    bool default_row_zeroed;

    XlsxSheetFormatProps sheetFormatProps;

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

    XlsxPrintOptions printOptions;
    XlsxPageMargins pageMargins;
    XlsxPageSetup pageSetup;
    QRegularExpression urlPattern;
private:
    static double calculateColWidth(int characters);
};

}
#endif // XLSXWORKSHEET_P_H
