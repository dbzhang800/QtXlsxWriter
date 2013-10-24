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
#include "xlsxworksheet.h"
#include "xlsxworksheet_p.h"
#include "xlsxworkbook.h"
#include "xlsxformat.h"
#include "xlsxutility_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxxmlwriter_p.h"
#include "xlsxxmlreader_p.h"
#include "xlsxdrawing_p.h"
#include "xlsxstyles_p.h"
#include "xlsxcell.h"
#include "xlsxcell_p.h"

#include <QVariant>
#include <QDateTime>
#include <QPoint>
#include <QFile>
#include <QUrl>
#include <QRegularExpression>
#include <QDebug>
#include <QBuffer>

#include <stdint.h>

QT_BEGIN_NAMESPACE_XLSX

WorksheetPrivate::WorksheetPrivate(Worksheet *p) :
    q_ptr(p)
{
    drawing = 0;

    xls_rowmax = 1048576;
    xls_colmax = 16384;
    xls_strmax = 32767;
    dim_rowmin = INT32_MAX;
    dim_rowmax = INT32_MIN;
    dim_colmin = INT32_MAX;
    dim_colmax = INT32_MIN;

    previous_row = 0;

    outline_row_level = 0;
    outline_col_level = 0;

    default_row_height = 15;
    default_row_zeroed = false;

    hidden = false;
    selected = false;
    right_to_left = false;
    show_zeros = true;
}

WorksheetPrivate::~WorksheetPrivate()
{
    if (drawing)
        delete drawing;
}

/*
  Calculate the "spans" attribute of the <row> tag. This is an
  XLSX optimisation and isn't strictly required. However, it
  makes comparing files easier. The span is the same for each
  block of 16 rows.
 */
void WorksheetPrivate::calculateSpans()
{
    row_spans.clear();
    int span_min = INT32_MAX;
    int span_max = INT32_MIN;

    for (int row_num = dim_rowmin; row_num <= dim_rowmax; row_num++) {
        if (cellTable.contains(row_num)) {
            for (int col_num = dim_colmin; col_num <= dim_colmax; col_num++) {
                if (cellTable[row_num].contains(col_num)) {
                    if (span_max == INT32_MIN) {
                        span_min = col_num;
                        span_max = col_num;
                    } else {
                        if (col_num < span_min)
                            span_min = col_num;
                        if (col_num > span_max)
                            span_max = col_num;
                    }
                }
            }
        }
        if (comments.contains(row_num)) {
            for (int col_num = dim_colmin; col_num <= dim_colmax; col_num++) {
                if (comments[row_num].contains(col_num)) {
                    if (span_max == INT32_MIN) {
                        span_min = col_num;
                        span_max = col_num;
                    } else {
                        if (col_num < span_min)
                            span_min = col_num;
                        if (col_num > span_max)
                            span_max = col_num;
                    }
                }
            }
        }

        if ((row_num + 1)%16 == 0 || row_num == dim_rowmax) {
            int span_index = row_num / 16;
            if (span_max != INT32_MIN) {
                span_min += 1;
                span_max += 1;
                row_spans[span_index] = QStringLiteral("%1:%2").arg(span_min).arg(span_max);
                span_max = INT32_MIN;
            }
        }
    }
}


QString WorksheetPrivate::generateDimensionString()
{
    if (dim_rowmax == INT32_MIN && dim_colmax == INT32_MIN) {
        //If the max dimensions are equal to INT32_MIN, then no dimension have been set
        //and we use the default "A1"
        return QStringLiteral("A1");
    }

    if (dim_rowmax == INT32_MIN) {
        //row dimensions aren't set but the column dimensions are set
        if (dim_colmin == dim_colmax) {
            //The dimensions are a single cell and not a range
            return xl_rowcol_to_cell(0, dim_colmin);
        } else {
            const QString cell_1 = xl_rowcol_to_cell(0, dim_colmin);
            const QString cell_2 = xl_rowcol_to_cell(0, dim_colmax);
            return cell_1 + QLatin1String(":") + cell_2;
        }
    }

    if (dim_rowmin == dim_rowmax && dim_colmin == dim_colmax) {
        //Single cell
        return xl_rowcol_to_cell(dim_rowmin, dim_rowmin);
    }

    QString cell_1 = xl_rowcol_to_cell(dim_rowmin, dim_colmin);
    QString cell_2 = xl_rowcol_to_cell(dim_rowmax, dim_colmax);
    return cell_1 + QLatin1String(":") + cell_2;
}

/*
  Check that row and col are valid and store the max and min
  values for use in other methods/elements. The ignore_row /
  ignore_col flags is used to indicate that we wish to perform
  the dimension check without storing the value. The ignore
  flags are use by setRow() and dataValidate.
*/
int WorksheetPrivate::checkDimensions(int row, int col, bool ignore_row, bool ignore_col)
{
    if (row >= xls_rowmax || col >= xls_colmax)
        return -1;

    if (!ignore_row) {
        if (row < dim_rowmin) dim_rowmin = row;
        if (row > dim_rowmax) dim_rowmax = row;
    }
    if (!ignore_col) {
        if (col < dim_colmin) dim_colmin = col;
        if (col > dim_colmax) dim_colmax = col;
    }

    return 0;
}

/*!
 * \brief Worksheet::Worksheet
 * \param name Name of the worksheet
 * \param index Index of the worksheet in the workbook
 * \param parent
 */
Worksheet::Worksheet(const QString &name, Workbook *workbook) :
    d_ptr(new WorksheetPrivate(this))
{
    d_ptr->name = name;
    if (!workbook) //For unit test propose only. Ignore the memery leak.
        workbook = new Workbook;
    d_ptr->workbook = workbook;
}

Worksheet::~Worksheet()
{
    delete d_ptr;
}

bool Worksheet::isChartsheet() const
{
    return false;
}

QString Worksheet::sheetName() const
{
    Q_D(const Worksheet);
    return d->name;
}

void Worksheet::setSheetName(const QString &sheetName)
{
    Q_D(Worksheet);
    d->name = sheetName;
}

bool Worksheet::isHidden() const
{
    Q_D(const Worksheet);
    return d->hidden;
}

bool Worksheet::isSelected() const
{
    Q_D(const Worksheet);
    return d->selected;
}

void Worksheet::setHidden(bool hidden)
{
    Q_D(Worksheet);
    d->hidden = hidden;
}

void Worksheet::setSelected(bool select)
{
    Q_D(Worksheet);
    d->selected = select;
}

void Worksheet::setRightToLeft(bool enable)
{
    Q_D(Worksheet);
    d->right_to_left = enable;
}

void Worksheet::setZeroValuesHidden(bool enable)
{
    Q_D(Worksheet);
    d->show_zeros = !enable;
}

QStringList Worksheet::externUrlList() const
{
    Q_D(const Worksheet);
    return d->externUrlList;
}

QStringList Worksheet::externDrawingList() const
{
    Q_D(const Worksheet);
    return d->externDrawingList;
}

QList<QPair<QString, QString> > Worksheet::drawingLinks() const
{
    Q_D(const Worksheet);
    return d->drawingLinks;
}

int Worksheet::write(int row, int column, const QVariant &value, Format *format)
{
    Q_D(Worksheet);
    bool ok;
    int ret = 0;

    if (d->checkDimensions(row, column))
        return -1;

    if (value.isNull()) { //blank
        ret = writeBlank(row, column, format);
    } else if (value.type() == QMetaType::Bool) { //Bool
        ret = writeBool(row,column, value.toBool(), format);
    } else if (value.toDateTime().isValid()) { //DateTime
        ret = writeDateTime(row, column, value.toDateTime(), format);
    } else if (value.toDouble(&ok), ok) { //Number
        if (!d->workbook->isStringsToNumbersEnabled() && value.type() == QMetaType::QString) {
            //Don't convert string to number if the flag not enabled.
            ret = writeString(row, column, value.toString(), format);
        } else {
            ret = writeNumeric(row, column, value.toDouble(), format);
        }
    } else if (value.type() == QMetaType::QUrl) { //url
        ret = writeUrl(row, column, value.toUrl(), format);
    } else if (value.type() == QMetaType::QString) { //string
        QString token = value.toString();
        QRegularExpression urlPattern(QStringLiteral("^([fh]tt?ps?://)|(mailto:)|((in|ex)ternal:)"));
        if (token.startsWith(QLatin1String("="))) {
            ret = writeFormula(row, column, token, format);
        } else if (token.startsWith(QLatin1String("{")) && token.endsWith(QLatin1String("}"))) {

        } else if (token.contains(urlPattern)) {
            ret = writeUrl(row, column, QUrl(token));
        } else {
            ret = writeString(row, column, token, format);
        }
    } else { //Wrong type

        return -1;
    }

    return ret;
}

//convert the "A1" notation to row/column notation
int Worksheet::write(const QString row_column, const QVariant &value, Format *format)
{
    QPoint pos = xl_cell_to_rowcol(row_column);
    if (pos == QPoint(-1, -1)) {
        return -1;
    }
    return write(pos.x(), pos.y(), value, format);
}

Cell *Worksheet::cellAt(const QString &row_column) const
{
    QPoint pos = xl_cell_to_rowcol(row_column);
    if (pos == QPoint(-1, -1))
        return 0;

    return cellAt(pos.x(), pos.y());
}

Cell *Worksheet::cellAt(int row, int column) const
{
    Q_D(const Worksheet);
    if (!d->cellTable.contains(row))
        return 0;
    if (!d->cellTable[row].contains(column))
        return 0;

    return d->cellTable[row][column].data();
}

int Worksheet::writeString(int row, int column, const QString &value, Format *format)
{
    Q_D(Worksheet);
    int error = 0;
    QString content = value;
    if (d->checkDimensions(row, column))
        return -1;

    if (value.size() > d->xls_strmax) {
        content = value.left(d->xls_strmax);
        error = -2;
    }

    SharedStrings *sharedStrings = d->workbook->sharedStrings();
    int index = sharedStrings->addSharedString(content);

    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(index, Cell::String, format));
    d->workbook->styles()->addFormat(format);
    return error;
}

int Worksheet::writeInlineString(int row, int column, const QString &value, Format *format)
{
    Q_D(Worksheet);
    int error = 0;
    QString content = value;
    if (d->checkDimensions(row, column))
        return -1;

    if (value.size() > d->xls_strmax) {
        content = value.left(d->xls_strmax);
        error = -2;
    }

    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::InlineString, format));
    d->workbook->styles()->addFormat(format);
    return error;
}

int Worksheet::writeNumeric(int row, int column, double value, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::Numeric, format));
    d->workbook->styles()->addFormat(format);
    return 0;
}

int Worksheet::writeFormula(int row, int column, const QString &content, Format *format, double result)
{
    Q_D(Worksheet);
    int error = 0;
    QString formula = content;
    if (d->checkDimensions(row, column))
        return -1;

    //Remove the formula '=' sign if exists
    if (formula.startsWith(QLatin1String("=")))
        formula.remove(0,1);

    Cell *data = new Cell(result, Cell::Formula, format);
    data->d_ptr->formula = formula;
    d->cellTable[row][column] = QSharedPointer<Cell>(data);
    d->workbook->styles()->addFormat(format);

    return error;
}

int Worksheet::writeBlank(int row, int column, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(QVariant(), Cell::Blank, format));
    d->workbook->styles()->addFormat(format);

    return 0;
}

int Worksheet::writeBool(int row, int column, bool value, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::Boolean, format));
    d->workbook->styles()->addFormat(format);

    return 0;
}

int Worksheet::writeDateTime(int row, int column, const QDateTime &dt, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    if (!format) {
        format = d->workbook->createFormat();
        format->setNumberFormat(d->workbook->defaultDateFormat());
    }

    double value = datetimeToNumber(dt, d->workbook->isDate1904());

    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::Numeric, format));
    d->workbook->styles()->addFormat(format);

    return 0;
}

int Worksheet::writeUrl(int row, int column, const QUrl &url, Format *format, const QString &display, const QString &tip)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    int link_type = 1;
    QString urlString = url.toString();
    QString displayString = display;
    if (urlString.startsWith(QLatin1String("internal:"))) {
        urlString.replace(QLatin1String("internal:"), QString());
        link_type = 2;
    } else if (urlString.startsWith(QLatin1String("external:"))) {
        urlString.replace(QLatin1String("external:"), QString());
        link_type = 3;
    }

    if (display.isEmpty())
        displayString = urlString;

    //For external links, chagne the directory separator from Unix to Dos
    if (link_type == 3) {
        urlString.replace(QLatin1Char('/'), QLatin1String("\\"));
        displayString.replace(QLatin1Char('/'), QLatin1String("\\"));
    }

    displayString.replace(QLatin1String("mailto:"), QString());

    int error = 0;
    if (displayString.size() > d->xls_strmax) {
        displayString = displayString.left(d->xls_strmax);
        error = -2;
    }

    QString locationString = displayString;
    if (link_type == 1) {
        locationString = QString();
    } else if (link_type == 3) {
        // External Workbook links need to be modified into correct format.
        // The URL will look something like 'c:\temp\file.xlsx#Sheet!A1'.
        // We need the part to the left of the # as the URL and the part to
        //the right as the "location" string (if it exists).
        if (urlString.contains(QLatin1Char('#'))) {
            QStringList list = urlString.split(QLatin1Char('#'));
            urlString = list[0];
            locationString = list[1];
        } else {
            locationString = QString();
        }
        link_type = 1;
    }


    //Write the hyperlink string as normal string.
    SharedStrings *sharedStrings = d->workbook->sharedStrings();
    int index = sharedStrings->addSharedString(urlString);
    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(index, Cell::String, format));

    //Store the hyperlink data in sa separate table
    d->urlTable[row][column] = new XlsxUrlData(link_type, urlString, locationString, tip);
    d->workbook->styles()->addFormat(format);

    return error;
}

int Worksheet::insertImage(int row, int column, const QImage &image, const QPointF &offset, double xScale, double yScale)
{
    Q_D(Worksheet);

    d->imageList.append(new XlsxImageData(row, column, image, offset, xScale, yScale));
    return 0;
}

int Worksheet::mergeCells(const QString &range)
{
    Q_D(Worksheet);
    QStringList cells = range.split(QLatin1Char(':'));
    if (cells.size() != 2)
        return -1;
    QPoint cell1 = xl_cell_to_rowcol(cells[0]);
    QPoint cell2 = xl_cell_to_rowcol(cells[1]);

    if (cell1 == QPoint(-1,-1) || cell2 == QPoint(-1, -1))
        return -1;

    return mergeCells(cell1.x(), cell1.y(), cell2.x(), cell2.y());
}

int Worksheet::mergeCells(int row_begin, int column_begin, int row_end, int column_end)
{
    Q_D(Worksheet);

    if (row_begin == row_end && column_begin == column_end)
        return -1;

    if (d->checkDimensions(row_end, column_end))
        return -1;

    XlsxCellRange range;
    range.row_begin = row_begin;
    range.row_end = row_end;
    range.column_begin = column_begin;
    range.column_end = column_end;

    d->merges.append(range);

    return 0;
}

int Worksheet::unmergeCells(const QString &range)
{
    Q_D(Worksheet);
    QStringList cells = range.split(QLatin1Char(':'));
    if (cells.size() != 2)
        return -1;
    QPoint cell1 = xl_cell_to_rowcol(cells[0]);
    QPoint cell2 = xl_cell_to_rowcol(cells[1]);

    if (cell1 == QPoint(-1,-1) || cell2 == QPoint(-1, -1))
        return -1;

    return unmergeCells(cell1.x(), cell1.y(), cell2.x(), cell2.y());
}

int Worksheet::unmergeCells(int row_begin, int column_begin, int row_end, int column_end)
{
    Q_D(Worksheet);
    XlsxCellRange range;
    range.row_begin = row_begin;
    range.row_end = row_end;
    range.column_begin = column_begin;
    range.column_end = column_end;

    if (!d->merges.contains(range))
        return -1;

    d->merges.removeOne(range);

    return 0;
}

void Worksheet::saveToXmlFile(QIODevice *device)
{
    Q_D(Worksheet);
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("worksheet"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    //for Excel 2010
    //    writer.writeAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    //    writer.writeAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
    //    writer.writeAttribute("mc:Ignorable", "x14ac");

    writer.writeStartElement(QStringLiteral("dimension"));
    writer.writeAttribute(QStringLiteral("ref"), d->generateDimensionString());
    writer.writeEndElement();//dimension

    writer.writeStartElement(QStringLiteral("sheetViews"));
    writer.writeStartElement(QStringLiteral("sheetView"));
    if (!d->show_zeros)
        writer.writeAttribute(QStringLiteral("showZeros"), QStringLiteral("0"));
    if (d->right_to_left)
        writer.writeAttribute(QStringLiteral("rightToLeft"), QStringLiteral("1"));
    if (d->selected)
        writer.writeAttribute(QStringLiteral("tabSelected"), QStringLiteral("1"));
    writer.writeAttribute(QStringLiteral("workbookViewId"), QStringLiteral("0"));
    writer.writeEndElement();//sheetView
    writer.writeEndElement();//sheetViews

    writer.writeStartElement(QStringLiteral("sheetFormatPr"));
    writer.writeAttribute(QStringLiteral("defaultRowHeight"), QString::number(d->default_row_height));
    if (d->default_row_height != 15)
        writer.writeAttribute(QStringLiteral("customHeight"), QStringLiteral("1"));
    if (d->default_row_zeroed)
        writer.writeAttribute(QStringLiteral("zeroHeight"), QStringLiteral("1"));
    if (d->outline_row_level)
        writer.writeAttribute(QStringLiteral("outlineLevelRow"), QString::number(d->outline_row_level));
    if (d->outline_col_level)
        writer.writeAttribute(QStringLiteral("outlineLevelCol"), QString::number(d->outline_col_level));
    //for Excel 2010
    //    writer.writeAttribute("x14ac:dyDescent", "0.25");
    writer.writeEndElement();//sheetFormatPr

    if (!d->colsInfo.isEmpty()) {
        writer.writeStartElement(QStringLiteral("cols"));
        for (int i=0; i<d->colsInfo.size(); ++i) {
            QSharedPointer<XlsxColumnInfo> col_info = d->colsInfo[i];
            writer.writeStartElement(QStringLiteral("col"));
            writer.writeAttribute(QStringLiteral("min"), QString::number(col_info->column_min + 1));
            writer.writeAttribute(QStringLiteral("max"), QString::number(col_info->column_max));
            writer.writeAttribute(QStringLiteral("width"), QString::number(col_info->width, 'g', 15));
            if (col_info->format)
                writer.writeAttribute(QStringLiteral("style"), QString::number(col_info->format->xfIndex()));
            if (col_info->hidden)
                writer.writeAttribute(QStringLiteral("hidden"), QStringLiteral("1"));
            if (col_info->width)
                writer.writeAttribute(QStringLiteral("customWidth"), QStringLiteral("1"));
            writer.writeEndElement();//col
        }
        writer.writeEndElement();//cols
    }

    writer.writeStartElement(QStringLiteral("sheetData"));
    if (d->dim_rowmax == INT32_MIN) {
        //If the max dimensions are equal to INT32_MIN, then there is no data to write
    } else {
        d->writeSheetData(writer);
    }
    writer.writeEndElement();//sheetData

    d->writeMergeCells(writer);
    d->writeHyperlinks(writer);
    d->writeDrawings(writer);

    writer.writeEndElement();//worksheet
    writer.writeEndDocument();
}

void WorksheetPrivate::writeSheetData(XmlStreamWriter &writer)
{
    calculateSpans();
    for (int row_num = dim_rowmin; row_num <= dim_rowmax; row_num++) {
        if (!(cellTable.contains(row_num) || comments.contains(row_num) || rowsInfo.contains(row_num))) {
            //Only process rows with cell data / comments / formatting
            continue;
        }

        int span_index = row_num / 16;
        QString span;
        if (row_spans.contains(span_index))
            span = row_spans[span_index];

        if (cellTable.contains(row_num)) {
            writer.writeStartElement(QStringLiteral("row"));
            writer.writeAttribute(QStringLiteral("r"), QString::number(row_num + 1));

            if (!span.isEmpty())
                writer.writeAttribute(QStringLiteral("spans"), span);

            if (rowsInfo.contains(row_num)) {
                QSharedPointer<XlsxRowInfo> rowInfo = rowsInfo[row_num];
                if (rowInfo->format) {
                    writer.writeAttribute(QStringLiteral("s"), QString::number(rowInfo->format->xfIndex()));
                    writer.writeAttribute(QStringLiteral("customFormat"), QStringLiteral("1"));
                }
                if (rowInfo->height != 15) {
                    writer.writeAttribute(QStringLiteral("ht"), QString::number(rowInfo->height));
                    writer.writeAttribute(QStringLiteral("customHeight"), QStringLiteral("1"));
                }
                if (rowInfo->hidden)
                    writer.writeAttribute(QStringLiteral("hidden"), QStringLiteral("1"));
            }

            for (int col_num = dim_colmin; col_num <= dim_colmax; col_num++) {
                if (cellTable[row_num].contains(col_num)) {
                    writeCellData(writer, row_num, col_num, cellTable[row_num][col_num]);
                }
            }
            writer.writeEndElement(); //row
        } else if (comments.contains(row_num)){

        } else {

        }
    }
}

void WorksheetPrivate::writeCellData(XmlStreamWriter &writer, int row, int col, QSharedPointer<Cell> cell)
{
    //This is the innermost loop so efficiency is important.
    QString cell_range = xl_rowcol_to_cell_fast(row, col);

    writer.writeStartElement(QStringLiteral("c"));
    writer.writeAttribute(QStringLiteral("r"), cell_range);

    //Style used by the cell, row or col
    if (cell->format())
        writer.writeAttribute(QStringLiteral("s"), QString::number(cell->format()->xfIndex()));
    else if (rowsInfo.contains(row) && rowsInfo[row]->format)
        writer.writeAttribute(QStringLiteral("s"), QString::number(rowsInfo[row]->format->xfIndex()));
    else if (colsInfoHelper.contains(col) && colsInfoHelper[col]->format)
        writer.writeAttribute(QStringLiteral("s"), QString::number(colsInfoHelper[col]->format->xfIndex()));

    if (cell->dataType() == Cell::String) {
        //cell->data: Index of the string in sharedStringTable
        writer.writeAttribute(QStringLiteral("t"), QStringLiteral("s"));
        writer.writeTextElement(QStringLiteral("v"), cell->value().toString());
    } else if (cell->dataType() == Cell::InlineString) {
        writer.writeAttribute(QStringLiteral("t"), QStringLiteral("inlineStr"));
        writer.writeStartElement(QStringLiteral("is"));
        writer.writeTextElement(QStringLiteral("t"), cell->value().toString());
        writer.writeEndElement();//is
    } else if (cell->dataType() == Cell::Numeric){
        double value = cell->value().toDouble();
        writer.writeTextElement(QStringLiteral("v"), QString::number(value, 'g', 15));
    } else if (cell->dataType() == Cell::Formula) {
        bool ok = true;
        cell->formula().toDouble(&ok);
        if (!ok) //is string
            writer.writeAttribute(QStringLiteral("t"), QStringLiteral("str"));
        writer.writeTextElement(QStringLiteral("f"), cell->formula());
        writer.writeTextElement(QStringLiteral("v"), cell->value().toString());
    } else if (cell->dataType() == Cell::Boolean) {
        writer.writeAttribute(QStringLiteral("t"), QStringLiteral("b"));
        writer.writeTextElement(QStringLiteral("v"), cell->value().toBool() ? QStringLiteral("1") : QStringLiteral("0"));
    } else if (cell->dataType() == Cell::Blank) {
        //Ok, empty here.
    }
    writer.writeEndElement(); //c
}

void WorksheetPrivate::writeMergeCells(XmlStreamWriter &writer)
{
    if (merges.isEmpty())
        return;

    writer.writeStartElement(QStringLiteral("mergeCells"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(merges.size()));

    foreach (XlsxCellRange range, merges) {
        QString cell1 = xl_rowcol_to_cell(range.row_begin, range.column_begin);
        QString cell2 = xl_rowcol_to_cell(range.row_end, range.column_end);
        writer.writeEmptyElement(QStringLiteral("mergeCell"));
        writer.writeAttribute(QStringLiteral("ref"), cell1+QLatin1Char(':')+cell2);
    }

    writer.writeEndElement(); //mergeCells
}

void WorksheetPrivate::writeHyperlinks(XmlStreamWriter &writer)
{
    if (urlTable.isEmpty())
        return;

    int rel_count = 0;
    externUrlList.clear();

    writer.writeStartElement(QStringLiteral("hyperlinks"));
    QMapIterator<int, QMap<int, XlsxUrlData *> > it(urlTable);
    while (it.hasNext()) {
        it.next();
        int row = it.key();
        QMapIterator <int, XlsxUrlData *> it2(it.value());
        while(it2.hasNext()) {
            it2.next();
            int col = it2.key();
            XlsxUrlData *data = it2.value();
            QString ref = xl_rowcol_to_cell(row, col);
            writer.writeEmptyElement(QStringLiteral("hyperlink"));
            writer.writeAttribute(QStringLiteral("ref"), ref);
            if (data->linkType == 1) {
                rel_count += 1;
                externUrlList.append(data->url);
                writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(rel_count));
                if (!data->location.isEmpty())
                    writer.writeAttribute(QStringLiteral("location"), data->location);
//                if (!data->url.isEmpty())
//                    writer.writeAttribute(QStringLiteral("display"), data->url);
                if (!data->tip.isEmpty())
                    writer.writeAttribute(QStringLiteral("tooltip"), data->tip);
            } else {
                writer.writeAttribute(QStringLiteral("location"), data->url);
                if (!data->tip.isEmpty())
                    writer.writeAttribute(QStringLiteral("tooltip"), data->tip);
                writer.writeAttribute(QStringLiteral("display"), data->location);
            }
        }
    }

    writer.writeEndElement();//hyperlinks
}

void WorksheetPrivate::writeDrawings(XmlStreamWriter &writer)
{
    if (!drawing)
        return;

    int index = externUrlList.size() + 1;
    writer.writeEmptyElement(QStringLiteral("drawing"));
    writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(index));
}

/*
  Sets row height and format. Row height measured in point size. If format
  equals 0 then format is ignored.
 */
bool Worksheet::setRow(int row, double height, Format *format, bool hidden)
{
    Q_D(Worksheet);
    int min_col = d->dim_colmax == INT32_MIN ? 0 : d->dim_colmin;

    if (d->checkDimensions(row, min_col))
        return false;

    d->rowsInfo[row] = QSharedPointer<XlsxRowInfo>(new XlsxRowInfo(height, format, hidden));
    d->workbook->styles()->addFormat(format);
    return true;
}

/*
  Sets column width and format for all columns from colFirst to colLast. Column
  width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. If format
  equals 0 then format is ignored.
 */
bool Worksheet::setColumn(int colFirst, int colLast, double width, Format *format, bool hidden)
{
    Q_D(Worksheet);
    bool ignore_row = true;
    bool ignore_col = (format || (width && hidden)) ? false : true;

    if (colFirst >= colLast)
        return false;

    if (d->checkDimensions(0, colLast, ignore_row, ignore_col))
        return false;
    if (d->checkDimensions(0, colFirst, ignore_row, ignore_col))
        return false;

    QSharedPointer<XlsxColumnInfo> info(new XlsxColumnInfo(colFirst, colLast, width, format, hidden));
    d->colsInfo.append(info);

    for (int col=colFirst; col<colLast; ++col)
        d->colsInfoHelper[col] = info;

    d->workbook->styles()->addFormat(format);

    return true;
}

Drawing *Worksheet::drawing() const
{
    Q_D(const Worksheet);
    return d->drawing;
}

QList<XlsxImageData *> Worksheet::images() const
{
    Q_D(const Worksheet);
    return d->imageList;
}

void Worksheet::clearExtraDrawingInfo()
{
    Q_D(Worksheet);
    if (d->drawing) {
        delete d->drawing;
        d->drawing = 0;
        d->externDrawingList.clear();
        d->drawingLinks.clear();
    }
}

void Worksheet::prepareImage(int index, int image_id, int drawing_id)
{
    Q_D(Worksheet);
    if (!d->drawing) {
        d->drawing = new Drawing;
        d->drawing->embedded = true;
        d->externDrawingList.append(QStringLiteral("../drawings/drawing%1.xml").arg(drawing_id));
    }

    XlsxImageData *imageData = d->imageList[index];

    XlsxDrawingDimensionData *data = new XlsxDrawingDimensionData;
    data->drawing_type = 2;

    double width = imageData->image.width() * imageData->xScale;
    double height = imageData->image.height() * imageData->yScale;

    XlsxObjectPositionData posData = d->pixelsToEMUs(d->objectPixelsPosition(imageData->col, imageData->row, imageData->offset.x(), imageData->offset.y(), width, height));
    data->col_from = posData.col_start;
    data->col_from_offset = posData.x1;
    data->row_from = posData.row_start;
    data->row_from_offset = posData.y1;
    data->col_to = posData.col_end;
    data->col_to_offset = posData.x2;
    data->row_to = posData.row_end;
    data->row_to_offset = posData.y2;
    data->width = posData.width;
    data->height = posData.height;
    data->col_absolute = posData.x_abs;
    data->row_absolute = posData.y_abs;

    d->drawing->dimensionList.append(data);

    d->drawingLinks.append(QPair<QString, QString>(QStringLiteral("/image"), QStringLiteral("../media/image%1.png").arg(image_id)));
}

/*
 Convert the height of a cell from user's units to pixels. If the
 height hasn't been set by the user we use the default value. If
 the row is hidden it has a value of zero.
*/
int WorksheetPrivate::rowPixelsSize(int row)
{
    double height;
    if (row_sizes.contains(row))
        height = row_sizes[row];
    else
        height = default_row_height;
    return static_cast<int>(4.0 / 3.0 *height);
}

/*
 Convert the width of a cell from user's units to pixels. Excel rounds
 the column width to the nearest pixel. If the width hasn't been set
 by the user we use the default value. If the column is hidden it
 has a value of zero.
*/
int WorksheetPrivate::colPixelsSize(int col)
{
    double max_digit_width = 7.0; //For Calabri 11
    double padding = 5.0;
    int pixels = 0;

    if (col_sizes.contains(col)) {
        double width = col_sizes[col];
        if (width < 1)
            pixels = static_cast<int>(width * (max_digit_width + padding) + 0.5);
        else
            pixels = static_cast<int>(width * max_digit_width + 0.5) + padding;
    } else {
        pixels = 64;
    }
    return pixels;
}

/*
        col_start     Col containing upper left corner of object.
        x1            Distance to left side of object.
        row_start     Row containing top left corner of object.
        y1            Distance to top of object.
        col_end       Col containing lower right corner of object.
        x2            Distance to right side of object.
        row_end       Row containing bottom right corner of object.
        y2            Distance to bottom of object.
        width         Width of object frame.
        height        Height of object frame.
        x_abs         Absolute distance to left side of object.
        y_abs         Absolute distance to top side of object.
*/
XlsxObjectPositionData WorksheetPrivate::objectPixelsPosition(int col_start, int row_start, double x1, double y1, double width, double height)
{
    double x_abs = 0;
    double y_abs = 0;
    for (int col_id = 1; col_id < col_start; ++col_id)
        x_abs += colPixelsSize(col_id);
    x_abs += x1;
    for (int row_id = 1; row_id < row_start; ++row_id)
        y_abs += rowPixelsSize(row_id);
    y_abs += y1;

    // Adjust start column for offsets that are greater than the col width.
    while (x1 > colPixelsSize(col_start)) {
        x1 -= colPixelsSize(col_start);
        col_start += 1;
    }
    while (y1 > rowPixelsSize(row_start)) {
        y1 -= rowPixelsSize(row_start);
        row_start += 1;
    }

    int col_end = col_start;
    int row_end = row_start;
    double x2 = width + x1;
    double y2 = height + y1;

    while (x2 > colPixelsSize(col_end)) {
        x2 -= colPixelsSize(col_end);
        col_end += 1;
    }

    while (y2 > rowPixelsSize(row_end)) {
        y2 -= rowPixelsSize(row_end);
        row_end += 1;
    }

    XlsxObjectPositionData data;
    data.col_start = col_start;
    data.x1 = x1;
    data.row_start = row_start;
    data.y1 = y1;
    data.col_end = col_end;
    data.x2 = x2;
    data.row_end = row_end;
    data.y2 = y2;
    data.x_abs = x_abs;
    data.y_abs = y_abs;
    data.width = width;
    data.height = height;

    return data;
}

/*
        Calculate the vertices that define the position of a graphical
        object within the worksheet in EMUs.

        The vertices are expressed as English Metric Units (EMUs). There are
        12,700 EMUs per point. Therefore, 12,700 * 3 /4 = 9,525 EMUs per
        pixel
*/
XlsxObjectPositionData WorksheetPrivate::pixelsToEMUs(const XlsxObjectPositionData &data)
{
    XlsxObjectPositionData result = data;
    result.x1 = static_cast<int>(data.x1 * 9525 + 0.5);
    result.y1 = static_cast<int>(data.y1 * 9525 + 0.5);
    result.x2 = static_cast<int>(data.x2 * 9525 + 0.5);
    result.y2 = static_cast<int>(data.y2 * 9525 + 0.5);
    result.x_abs = static_cast<int>(data.x_abs * 9525 + 0.5);
    result.y_abs = static_cast<int>(data.y_abs * 9525 + 0.5);
    result.width = static_cast<int>(data.width * 9525 + 0.5);
    result.height = static_cast<int>(data.height * 9525 + 0.5);

    return result;
}

QByteArray Worksheet::saveToXmlData()
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);

    return data;
}

void WorksheetPrivate::readSheetData(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("sheetData"));

    while(!(reader.name() == QLatin1String("sheetData") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();

        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("row")) {
                QXmlStreamAttributes attributes = reader.attributes();

                if (attributes.hasAttribute(QLatin1String("customFormat"))
                        || attributes.hasAttribute(QLatin1String("customHeight"))
                        || attributes.hasAttribute(QLatin1String("hidden"))) {

                    QSharedPointer<XlsxRowInfo> info(new XlsxRowInfo);
                    if (attributes.hasAttribute(QLatin1String("customFormat")) && attributes.hasAttribute(QLatin1String("s"))) {
                        int idx = attributes.value(QLatin1String("s")).toInt();
                        info->format = workbook->styles()->xfFormat(idx);
                    }
                    if (attributes.hasAttribute(QLatin1String("customHeight")) && attributes.hasAttribute(QLatin1String("ht"))) {
                        info->height = attributes.value(QLatin1String("ht")).toDouble();
                    }
                    if (attributes.hasAttribute(QLatin1String("hidden")))
                        info->hidden = true;

                    int row = attributes.value(QLatin1String("r")).toInt();
                    rowsInfo[row] = info;
                }

            } else if (reader.name() == QLatin1String("c")) {
                QXmlStreamAttributes attributes = reader.attributes();
                QString r = attributes.value(QLatin1String("r")).toString();
                QPoint pos = xl_cell_to_rowcol(r);

                //get format
                Format *format = 0;
                if (attributes.hasAttribute(QLatin1String("s"))) {
                    int idx = attributes.value(QLatin1String("s")).toInt();
                    format = workbook->styles()->xfFormat(idx);
                }

                if (attributes.hasAttribute(QLatin1String("t"))) {
                    QString type = attributes.value(QLatin1String("t")).toString();
                    if (type == QLatin1String("s")) {
                        //string type
                        reader.readNextStartElement();
                        if (reader.name() == QLatin1String("v")) {
                            QString value = reader.readElementText();
                            workbook->sharedStrings()->incRefByStringIndex(value.toInt());
                            Cell *data = new Cell(value ,Cell::String, format);
                            cellTable[pos.x()][pos.y()] = QSharedPointer<Cell>(data);
                        }
                    } else if (type == QLatin1String("inlineStr")) {
                        //inline string type
                        while (!(reader.name() == QLatin1String("c") && reader.tokenType() == QXmlStreamReader::EndElement)) {
                            reader.readNextStartElement();
                            if (reader.tokenType() == QXmlStreamReader::StartElement) {
                                if (reader.name() == QLatin1String("t")) {
                                    QString value = reader.readElementText();
                                    QSharedPointer<Cell> data(new Cell(value, Cell::InlineString, format));
                                    cellTable[pos.x()][pos.y()] = data;
                                }
                            }
                        }
                    } else if (type == QLatin1String("b")) {
                        //bool type
                        reader.readNextStartElement();
                        if (reader.name() == QLatin1String("v")) {
                            QString value = reader.readElementText();
                            QSharedPointer<Cell> data(new Cell(value.toInt() ? true : false, Cell::Boolean, format));
                            cellTable[pos.x()][pos.y()] = data;
                        }
                    }
                } else {
                    //number type
                    reader.readNextStartElement();
                    if (reader.name() == QLatin1String("v")) {
                        QString value = reader.readElementText();
                        Cell *data = new Cell(value ,Cell::Numeric, format);
                        cellTable[pos.x()][pos.y()] = QSharedPointer<Cell>(data);
                    }
                }
            }
        }
    }
}

void WorksheetPrivate::readColumnsInfo(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("cols"));

    while(!(reader.name() == QLatin1String("cols") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("col")) {
                QSharedPointer<XlsxColumnInfo> info(new XlsxColumnInfo);

                QXmlStreamAttributes colAttrs = reader.attributes();
                int min = colAttrs.value(QLatin1String("min")).toInt();
                int max = colAttrs.value(QLatin1String("max")).toInt();
                info->column_min = min - 1;
                info->column_max = max;

                if (colAttrs.hasAttribute(QLatin1String("customWidth"))) {
                    double width = colAttrs.value(QLatin1String("width")).toDouble();
                    info->width = width;
                }

                if (colAttrs.hasAttribute(QLatin1String("hidden")))
                    info->hidden = true;

                if (colAttrs.hasAttribute(QLatin1String("style"))) {
                    int idx = colAttrs.value(QLatin1String("style")).toInt();
                    info->format = workbook->styles()->xfFormat(idx);
                }

                colsInfo.append(info);
                for (int col=min; col<max; ++col)
                    colsInfoHelper[col] = info;
            }
        }
    }
}

void WorksheetPrivate::readMergeCells(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("mergeCells"));

    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();

    while(!(reader.name() == QLatin1String("mergeCells") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("mergeCell")) {
                QXmlStreamAttributes attrs = reader.attributes();
                QString rangeStr = attrs.value(QLatin1String("ref")).toString();
                QStringList items = rangeStr.split(QLatin1Char(':'));
                if (items.size() != 2) {
                    //Error
                } else {
                    QPoint p0 = xl_cell_to_rowcol(items[0]);
                    QPoint p1 = xl_cell_to_rowcol(items[1]);

                    XlsxCellRange range;
                    range.row_begin = p0.x();
                    range.column_begin = p0.y();
                    range.row_end = p1.x();
                    range.column_end = p1.y();

                    merges.append(range);
                }
            }
        }
    }
}

bool Worksheet::loadFromXmlFile(QIODevice *device)
{
    Q_D(Worksheet);

    XmlStreamReader reader(device);
    while(!reader.atEnd()) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("dimension")) {
                QXmlStreamAttributes attributes = reader.attributes();
                QStringList range = attributes.value(QLatin1String("ref")).toString().split(QLatin1Char(':'));
                if (range.size() == 2) {
                    QPoint start = xl_cell_to_rowcol(range[0]);
                    QPoint end = xl_cell_to_rowcol(range[1]);
                    d->dim_rowmin = start.x();
                    d->dim_colmin = start.y();
                    d->dim_rowmax = end.x();
                    d->dim_colmax = end.y();
                } else {
                    QPoint p = xl_cell_to_rowcol(range[0]);
                    d->dim_rowmin = p.x();
                    d->dim_colmin = p.y();
                    d->dim_rowmax = p.x();
                    d->dim_colmax = p.y();
                }
            } else if (reader.name() == QLatin1String("sheetViews")) {

            } else if (reader.name() == QLatin1String("sheetFormatPr")) {

            } else if (reader.name() == QLatin1String("cols")) {
                d->readColumnsInfo(reader);
            } else if (reader.name() == QLatin1String("sheetData")) {
                d->readSheetData(reader);
            } else if (reader.name() == QLatin1String("mergeCells")) {
                d->readMergeCells(reader);
            }
        }
    }

    return true;
}

bool Worksheet::loadFromXmlData(const QByteArray &data)
{
    QBuffer buffer;
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);

    return loadFromXmlFile(&buffer);
}

QT_END_NAMESPACE_XLSX
