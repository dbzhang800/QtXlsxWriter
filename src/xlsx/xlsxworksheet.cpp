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

#include <QVariant>
#include <QDateTime>
#include <QPoint>
#include <QFile>
#include <QDebug>

#include <stdint.h>

namespace QXlsx {

WorksheetPrivate::WorksheetPrivate(Worksheet *p) :
    q_ptr(p)
{
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
    actived = false;
    right_to_left = false;
    show_zeros = true;
}

WorksheetPrivate::~WorksheetPrivate()
{
    typedef QMap<int, XlsxCellData *> RowMap;
    foreach (RowMap row, cellTable) {
        foreach (XlsxCellData *item, row)
            delete item;
    }

    foreach (XlsxRowInfo *row, rowsInfo)
        delete row;

    foreach (XlsxColumnInfo *col, colsInfo)
        delete col;
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
Worksheet::Worksheet(const QString &name, int index, Workbook *parent) :
    QObject(parent), d_ptr(new WorksheetPrivate(this))
{
    d_ptr->name = name;
    d_ptr->index = index;
    d_ptr->workbook = parent;
}

Worksheet::~Worksheet()
{
    delete d_ptr;
}

bool Worksheet::isChartsheet() const
{
    return false;
}

QString Worksheet::name() const
{
    Q_D(const Worksheet);
    return d->name;
}

int Worksheet::index() const
{
    Q_D(const Worksheet);
    return d->index;
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

bool Worksheet::isActived() const
{
    Q_D(const Worksheet);
    return d->actived;
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

void Worksheet::setActived(bool act)
{
    Q_D(Worksheet);
    d->actived = act;
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
            ret = writeNumber(row, column, value.toDouble(), format);
        }
    } else if (value.type() == QMetaType::QUrl) { //url

    } else if (value.type() == QMetaType::QString) { //string
        QString token = value.toString();
        if (token.startsWith(QLatin1String("="))) {
            ret = writeFormula(row, column, token, format);
        } else if (token.startsWith(QLatin1String("{")) && token.endsWith(QLatin1String("}"))) {

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

    d->cellTable[row][column] = new XlsxCellData(index, XlsxCellData::String, format);
    return error;
}

int Worksheet::writeNumber(int row, int column, double value, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    d->cellTable[row][column] = new XlsxCellData(value, XlsxCellData::Number, format);
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

    XlsxCellData *data = new XlsxCellData(result, XlsxCellData::Formula, format);
    data->formula = formula;
    d->cellTable[row][column] = data;

    return error;
}

int Worksheet::writeBlank(int row, int column, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    d->cellTable[row][column] = new XlsxCellData(QVariant(), XlsxCellData::Blank, format);
    return 0;
}

int Worksheet::writeBool(int row, int column, bool value, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    d->cellTable[row][column] = new XlsxCellData(value, XlsxCellData::Boolean, format);
    return 0;
}

int Worksheet::writeDateTime(int row, int column, const QDateTime &dt, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    d->cellTable[row][column] = new XlsxCellData(dt, XlsxCellData::DateTime, format);
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
        foreach (XlsxColumnInfo *col_info, d->colsInfo) {
            writer.writeStartElement(QStringLiteral("col"));
            writer.writeAttribute(QStringLiteral("min"), QString::number(col_info->column_min));
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
                XlsxRowInfo *rowInfo = rowsInfo[row_num];
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

void WorksheetPrivate::writeCellData(XmlStreamWriter &writer, int row, int col, XlsxCellData *cell)
{
    //This is the innermost loop so efficiency is important.
    QString cell_range = xl_rowcol_to_cell_fast(row, col);

    writer.writeStartElement(QStringLiteral("c"));
    writer.writeAttribute(QStringLiteral("r"), cell_range);

    //Style used by the cell, row or col
    if (cell->format)
        writer.writeAttribute(QStringLiteral("s"), QString::number(cell->format->xfIndex()));
    else if (rowsInfo.contains(row) && rowsInfo[row]->format)
        writer.writeAttribute(QStringLiteral("s"), QString::number(rowsInfo[row]->format->xfIndex()));
    else if (colsInfoHelper.contains(col) && colsInfoHelper[col]->format)
        writer.writeAttribute(QStringLiteral("s"), QString::number(colsInfoHelper[col]->format->xfIndex()));

    if (cell->dataType == XlsxCellData::String) {
        //cell->data: Index of the string in sharedStringTable
        writer.writeAttribute(QStringLiteral("t"), QStringLiteral("s"));
        writer.writeTextElement(QStringLiteral("v"), cell->value.toString());
    } else if (cell->dataType == XlsxCellData::Number){
        double value = cell->value.toDouble();
        writer.writeTextElement(QStringLiteral("v"), QString::number(value, 'g', 15));
    } else if (cell->dataType == XlsxCellData::Formula) {
        bool ok = true;
        cell->formula.toDouble(&ok);
        if (!ok) //is string
            writer.writeAttribute(QStringLiteral("t"), QStringLiteral("str"));
        writer.writeTextElement(QStringLiteral("f"), cell->formula);
        writer.writeTextElement(QStringLiteral("v"), cell->value.toString());
    } else if (cell->dataType == XlsxCellData::ArrayFormula) {

    } else if (cell->dataType == XlsxCellData::Boolean) {
        writer.writeAttribute(QStringLiteral("t"), QStringLiteral("b"));
        writer.writeTextElement(QStringLiteral("v"), cell->value.toBool() ? QStringLiteral("1") : QStringLiteral("0"));
    } else if (cell->dataType == XlsxCellData::Blank) {
        //Ok, empty here.
    } else if (cell->dataType == XlsxCellData::DateTime) {
        QDateTime epoch(QDate(1899, 12, 31));
        if (workbook->isDate1904())
            epoch = QDateTime(QDate(1904, 1, 1));
        qint64 delta = epoch.msecsTo(cell->value.toDateTime());
        double excel_time = delta / (1000*60*60*24);
        //Account for Excel erroneously treating 1900 as a leap year.
        if (!workbook->isDate1904() && excel_time > 59)
            excel_time += 1;
        writer.writeTextElement(QStringLiteral("v"), QString::number(excel_time, 'g', 15));
    }
    writer.writeEndElement(); //c
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

    if (d->rowsInfo.contains(row)) {
        d->rowsInfo[row]->height = height;
        d->rowsInfo[row]->format = format;
        d->rowsInfo[row]->hidden = hidden;
    } else {
        d->rowsInfo[row] = new XlsxRowInfo(height, format, hidden);
    }
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

    if (d->checkDimensions(0, colLast, ignore_row, ignore_col))
        return false;
    if (d->checkDimensions(0, colFirst, ignore_row, ignore_col))
        return false;

    XlsxColumnInfo *info = new XlsxColumnInfo(colFirst, colLast, width, format, hidden);
    d->colsInfo.append(info);

    for (int col=colFirst; col<=colLast; ++col)
        d->colsInfoHelper[col] = info;

    return true;
}

} //namespace
