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
#include "xlsxworkbook.h"
#include "xlsxformat.h"
#include "xlsxutility_p.h"
#include "xlsxsharedstrings_p.h"
#include "xmlstreamwriter_p.h"

#include <QVariant>
#include <QDateTime>
#include <QPoint>
#include <QFile>
#include <QDebug>

#include <stdint.h>

namespace QXlsx {

struct XlsxCellData
{
    enum CellDataType {
        Blank,
        String,
        Number,
        Formula,
        ArrayFormula,
        Boolean
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

/*!
 * \brief Worksheet::Worksheet
 * \param name Name of the worksheet
 * \param index Index of the worksheet in the workbook
 * \param parent
 */
Worksheet::Worksheet(const QString &name, int index, Workbook *parent) :
    QObject(parent), m_workbook(parent), m_name(name), m_index(index)
{
    m_xls_rowmax = 1048576;
    m_xls_colmax = 16384;
    m_xls_strmax = 32767;
    m_dim_rowmin = INT32_MAX;
    m_dim_rowmax = INT32_MIN;
    m_dim_colmin = INT32_MAX;
    m_dim_colmax = INT32_MIN;

    m_previous_row = 0;

    m_outline_row_level = 0;
    m_outline_col_level = 0;

    m_default_row_height = 15;
    m_default_row_zeroed = false;

    m_hidden = false;
    m_selected = false;
    m_actived = false;
    m_right_to_left = false;
    m_show_zeros = true;
}

Worksheet::~Worksheet()
{
    typedef QMap<int, XlsxCellData *> RowMap;
    foreach (RowMap row, m_cellTable) {
        foreach (XlsxCellData *item, row)
            delete item;
    }

    foreach (XlsxRowInfo *row, m_rowsInfo)
        delete row;

    foreach (XlsxColumnInfo *col, m_colsInfo)
        delete col;
}

bool Worksheet::isChartsheet() const
{
    return false;
}

QString Worksheet::name() const
{
    return m_name;
}

int Worksheet::index() const
{
    return m_index;
}

bool Worksheet::isHidden() const
{
    return m_hidden;
}

bool Worksheet::isSelected() const
{
    return m_selected;
}

bool Worksheet::isActived() const
{
    return m_actived;
}

void Worksheet::setHidden(bool hidden)
{
    m_hidden = hidden;
}

void Worksheet::setSelected(bool select)
{
    m_selected = select;
}

void Worksheet::setActived(bool act)
{
    m_actived = act;
}

void Worksheet::setRightToLeft(bool enable)
{
    m_right_to_left = enable;
}

void Worksheet::setZeroValuesHidden(bool enable)
{
    m_show_zeros = !enable;
}

int Worksheet::write(int row, int column, const QVariant &value, Format *format)
{
    bool ok;
    int ret = 0;

    if (value.isNull()) { //blank
        ret = writeBlank(row, column, format);
    } else if (value.type() == QMetaType::Bool) { //Bool
        ret = writeBool(row,column, value.toBool(), format);
    } else if (value.toDateTime().isValid()) { //DateTime

    } else if (value.toDouble(&ok), ok) { //Number
        if (!m_workbook->isStringsToNumbersEnabled() && value.type() == QMetaType::QString) {
            //Don't convert string to number if the flag not enabled.
            ret = writeString(row, column, value.toString(), format);
        } else {
            ret = writeNumber(row, column, value.toDouble(), format);
        }
    } else if (value.type() == QMetaType::QUrl) { //url

    } else if (value.type() == QMetaType::QString) { //string
        QString token = value.toString();
        if (token.startsWith("=")) {
            ret = writeFormula(row, column, token, format);
        } else if (token.startsWith("{") && token.endsWith("}")) {

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
    int error = 0;
    QString content = value;
    if (checkDimensions(row, column))
        return -1;

    if (value.size() > m_xls_strmax) {
        content = value.left(m_xls_strmax);
        error = -2;
    }

    SharedStrings *sharedStrings = m_workbook->sharedStrings();
    int index = sharedStrings->addSharedString(content);

    m_cellTable[row][column] = new XlsxCellData(index, XlsxCellData::String, format);
    return error;
}

int Worksheet::writeNumber(int row, int column, double value, Format *format)
{
    if (checkDimensions(row, column))
        return -1;

    m_cellTable[row][column] = new XlsxCellData(value, XlsxCellData::Number, format);
    return 0;
}

int Worksheet::writeFormula(int row, int column, const QString &content, Format *format, double result)
{
    int error = 0;
    QString formula = content;
    if (checkDimensions(row, column))
        return -1;

    //Remove the formula '=' sign if exists
    if (formula.startsWith("="))
        formula.remove(0,1);

    XlsxCellData *data = new XlsxCellData(result, XlsxCellData::Formula, format);
    data->formula = formula;
    m_cellTable[row][column] = data;

    return error;
}

int Worksheet::writeBlank(int row, int column, Format *format)
{
    if (checkDimensions(row, column))
        return -1;

    m_cellTable[row][column] = new XlsxCellData(QVariant(), XlsxCellData::Blank, format);
    return 0;
}

int Worksheet::writeBool(int row, int column, bool value, Format *format)
{
    if (checkDimensions(row, column))
        return -1;

    m_cellTable[row][column] = new XlsxCellData(value, XlsxCellData::Boolean, format);
    return 0;
}

/*
  Check that row and col are valid and store the max and min
  values for use in other methods/elements. The ignore_row /
  ignore_col flags is used to indicate that we wish to perform
  the dimension check without storing the value. The ignore
  flags are use by setRow() and dataValidate.
*/
int Worksheet::checkDimensions(int row, int col, bool ignore_row, bool ignore_col)
{
    if (row >= m_xls_rowmax || col >= m_xls_colmax)
        return -1;

    if (!ignore_row) {
        if (row < m_dim_rowmin) m_dim_rowmin = row;
        if (row > m_dim_rowmax) m_dim_rowmax = row;
    }
    if (!ignore_col) {
        if (col < m_dim_colmin) m_dim_colmin = col;
        if (col > m_dim_colmax) m_dim_colmax = col;
    }

    return 0;
}

void Worksheet::saveToXmlFile(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument("1.0", true);
    writer.writeStartElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

    //for Excel 2010
    //    writer.writeAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    //    writer.writeAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
    //    writer.writeAttribute("mc:Ignorable", "x14ac");

    writer.writeStartElement("dimension");
    writer.writeAttribute("ref", generateDimensionString());
    writer.writeEndElement();//dimension

    writer.writeStartElement("sheetViews");
    writer.writeStartElement("sheetView");
    if (!m_show_zeros)
        writer.writeAttribute("showZeros", "0");
    if (m_right_to_left)
        writer.writeAttribute("rightToLeft", "1");
    if (m_selected)
        writer.writeAttribute("tabSelected", "1");
    writer.writeAttribute("workbookViewId", "0");
    writer.writeEndElement();//sheetView
    writer.writeEndElement();//sheetViews

    writer.writeStartElement("sheetFormatPr");
    writer.writeAttribute("defaultRowHeight", QString::number(m_default_row_height));
    if (m_default_row_height != 15)
        writer.writeAttribute("customHeight", "1");
    if (m_default_row_zeroed)
        writer.writeAttribute("zeroHeight", "1");
    if (m_outline_row_level)
        writer.writeAttribute("outlineLevelRow", QString::number(m_outline_row_level));
    if (m_outline_col_level)
        writer.writeAttribute("outlineLevelCol", QString::number(m_outline_col_level));
    //for Excel 2010
    //    writer.writeAttribute("x14ac:dyDescent", "0.25");
    writer.writeEndElement();//sheetFormatPr

    if (!m_colsInfo.isEmpty()) {
        writer.writeStartElement("cols");
        foreach (XlsxColumnInfo *col_info, m_colsInfo) {
            writer.writeStartElement("col");
            writer.writeAttribute("min", QString::number(col_info->column_min));
            writer.writeAttribute("max", QString::number(col_info->column_max));
            writer.writeAttribute("width", QString::number(col_info->width, 'g', 15));
            if (col_info->format)
                writer.writeAttribute("style", QString::number(col_info->format->xfIndex()));
            if (col_info->hidden)
                writer.writeAttribute("hidden", "1");
            if (col_info->width)
                writer.writeAttribute("customWidth", "1");
            writer.writeEndElement();//col
        }
        writer.writeEndElement();//cols
    }

    writer.writeStartElement("sheetData");
    if (m_dim_rowmax == INT32_MIN) {
        //If the max dimensions are equal to INT32_MIN, then there is no data to write
    } else {
        writeSheetData(writer);
    }
    writer.writeEndElement();//sheetData

    writer.writeEndElement();//worksheet
    writer.writeEndDocument();
}

QString Worksheet::generateDimensionString()
{
    if (m_dim_rowmax == INT32_MIN && m_dim_colmax == INT32_MIN) {
        //If the max dimensions are equal to INT32_MIN, then no dimension have been set
        //and we use the default "A1"
        return "A1";
    }

    if (m_dim_rowmax == INT32_MIN) {
        //row dimensions aren't set but the column dimensions are set
        if (m_dim_colmin == m_dim_colmax) {
            //The dimensions are a single cell and not a range
            return xl_rowcol_to_cell(0, m_dim_colmin);
        } else {
            const QString cell_1 = xl_rowcol_to_cell(0, m_dim_colmin);
            const QString cell_2 = xl_rowcol_to_cell(0, m_dim_colmax);
            return cell_1 + ":" + cell_2;
        }
    }

    if (m_dim_rowmin == m_dim_rowmax && m_dim_colmin == m_dim_colmax) {
        //Single cell
        return xl_rowcol_to_cell(m_dim_rowmin, m_dim_rowmin);
    }

    QString cell_1 = xl_rowcol_to_cell(m_dim_rowmin, m_dim_colmin);
    QString cell_2 = xl_rowcol_to_cell(m_dim_rowmax, m_dim_colmax);
    return cell_1 + ":" + cell_2;
}

void Worksheet::writeSheetData(XmlStreamWriter &writer)
{
    calculateSpans();
    for (int row_num = m_dim_rowmin; row_num <= m_dim_rowmax; row_num++) {
        if (!(m_cellTable.contains(row_num) || m_comments.contains(row_num) || m_rowsInfo.contains(row_num))) {
            //Only process rows with cell data / comments / formatting
            continue;
        }

        int span_index = row_num / 16;
        QString span;
        if (m_row_spans.contains(span_index))
            span = m_row_spans[span_index];

        if (m_cellTable.contains(row_num)) {
            writer.writeStartElement("row");
            writer.writeAttribute("r", QString::number(row_num + 1));

            if (!span.isEmpty())
                writer.writeAttribute("spans", span);

            if (m_rowsInfo.contains(row_num)) {
                XlsxRowInfo *rowInfo = m_rowsInfo[row_num];
                if (rowInfo->format) {
                    writer.writeAttribute("s", QString::number(rowInfo->format->xfIndex()));
                    writer.writeAttribute("customFormat", "1");
                }
                if (rowInfo->height != 15) {
                    writer.writeAttribute("ht", QString::number(rowInfo->height));
                    writer.writeAttribute("customHeight", "1");
                }
                if (rowInfo->hidden)
                    writer.writeAttribute("hidden", "1");
            }

            for (int col_num = m_dim_colmin; col_num <= m_dim_colmax; col_num++) {
                if (m_cellTable[row_num].contains(col_num)) {
                    writeCellData(writer, row_num, col_num, m_cellTable[row_num][col_num]);
                }
            }
            writer.writeEndElement(); //row
        } else if (m_comments.contains(row_num)){

        } else {

        }
    }
}

void Worksheet::writeCellData(XmlStreamWriter &writer, int row, int col, XlsxCellData *cell)
{
    //This is the innermost loop so efficiency is important.
    QString cell_range = xl_rowcol_to_cell_fast(row, col);

    writer.writeStartElement("c");
    writer.writeAttribute("r", cell_range);

    //Style used by the cell, row or col
    if (cell->format)
        writer.writeAttribute("s", QString::number(cell->format->xfIndex()));
    else if (m_rowsInfo.contains(row) && m_rowsInfo[row]->format)
        writer.writeAttribute("s", QString::number(m_rowsInfo[row]->format->xfIndex()));
    else if (m_colsInfoHelper.contains(col) && m_colsInfoHelper[col]->format)
        writer.writeAttribute("s", QString::number(m_colsInfoHelper[col]->format->xfIndex()));

    if (cell->dataType == XlsxCellData::String) {
        //cell->data: Index of the string in sharedStringTable
        writer.writeAttribute("t", "s");
        writer.writeTextElement("v", cell->value.toString());
    } else if (cell->dataType == XlsxCellData::Number){
        writer.writeTextElement("v", cell->value.toString());
    } else if (cell->dataType == XlsxCellData::Formula) {
        bool ok = true;
        cell->formula.toDouble(&ok);
        if (!ok) //is string
            writer.writeAttribute("t", "str");
        writer.writeTextElement("f", cell->formula);
        writer.writeTextElement("v", cell->value.toString());
    } else if (cell->dataType == XlsxCellData::ArrayFormula) {

    } else if (cell->dataType == XlsxCellData::Boolean) {
        writer.writeAttribute("t", "b");
        writer.writeTextElement("v", cell->value.toBool() ? "1" : "0");
    } else if (cell->dataType == XlsxCellData::Blank) {
        //Ok, empty here.
    }
    writer.writeEndElement(); //c
}

/*
  Calculate the "spans" attribute of the <row> tag. This is an
  XLSX optimisation and isn't strictly required. However, it
  makes comparing files easier. The span is the same for each
  block of 16 rows.
 */
void Worksheet::calculateSpans()
{
    m_row_spans.clear();
    int span_min = INT32_MAX;
    int span_max = INT32_MIN;

    for (int row_num = m_dim_rowmin; row_num <= m_dim_rowmax; row_num++) {
        if (m_cellTable.contains(row_num)) {
            for (int col_num = m_dim_colmin; col_num <= m_dim_colmax; col_num++) {
                if (m_cellTable[row_num].contains(col_num)) {
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
        if (m_comments.contains(row_num)) {
            for (int col_num = m_dim_colmin; col_num <= m_dim_colmax; col_num++) {
                if (m_comments[row_num].contains(col_num)) {
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

        if ((row_num + 1)%16 == 0 || row_num == m_dim_rowmax) {
            int span_index = row_num / 16;
            if (span_max != INT32_MIN) {
                span_min += 1;
                span_max += 1;
                m_row_spans[span_index] = QString("%1:%2").arg(span_min).arg(span_max);
                span_max = INT32_MIN;
            }
        }
    }
}

/*
  Sets row height and format. Row height measured in point size. If format
  equals 0 then format is ignored.
 */
bool Worksheet::setRow(int row, double height, Format *format, bool hidden)
{
    int min_col = m_dim_colmax == INT32_MIN ? 0 : m_dim_colmin;

    if (checkDimensions(row, min_col))
        return false;

    if (m_rowsInfo.contains(row)) {
        m_rowsInfo[row]->height = height;
        m_rowsInfo[row]->format = format;
        m_rowsInfo[row]->hidden = hidden;
    } else {
        m_rowsInfo[row] = new XlsxRowInfo(height, format, hidden);
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
    bool ignore_row = true;
    bool ignore_col = (format || (width && hidden)) ? false : true;

    if (checkDimensions(0, colLast, ignore_row, ignore_col))
        return false;
    if (checkDimensions(0, colFirst, ignore_row, ignore_col))
        return false;

    XlsxColumnInfo *info = new XlsxColumnInfo(colFirst, colLast, width, format, hidden);
    m_colsInfo.append(info);

    for (int col=colFirst; col<=colLast; ++col)
        m_colsInfoHelper[col] = info;

    return true;
}

} //namespace
