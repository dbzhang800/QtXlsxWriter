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
#include "xlsxcellrange.h"

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
  , windowProtection(false), showFormulas(false), showGridLines(true), showRowColHeaders(true)
  , showZeros(true), rightToLeft(false), tabSelected(false), showRuler(false)
  , showOutlineSymbols(true), showWhiteSpace(true)
{
    drawing = 0;

    xls_rowmax = 1048576;
    xls_colmax = 16384;
    xls_strmax = 32767;

    previous_row = 0;

    outline_row_level = 0;
    outline_col_level = 0;

    default_row_height = 15;
    default_row_zeroed = false;

    hidden = false;
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

    for (int row_num = dimension.firstRow(); row_num <= dimension.lastRow(); row_num++) {
        if (cellTable.contains(row_num)) {
            for (int col_num = dimension.firstColumn(); col_num <= dimension.lastColumn(); col_num++) {
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
            for (int col_num = dimension.firstColumn(); col_num <= dimension.lastColumn(); col_num++) {
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

        if ((row_num + 1)%16 == 0 || row_num == dimension.lastRow()) {
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
    if (!dimension.isValid())
        return QStringLiteral("A1");
    else
        return dimension.toString();
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
        if (row < dimension.firstRow() || dimension.firstRow() == -1) dimension.setFirstRow(row);
        if (row > dimension.lastRow()) dimension.setLastRow(row);
    }
    if (!ignore_col) {
        if (col < dimension.firstColumn() || dimension.firstColumn() == -1) dimension.setFirstColumn(col);
        if (col > dimension.lastColumn()) dimension.setLastColumn(col);
    }

    return 0;
}

/*!
 * \brief Worksheet::Worksheet
 * \param name Name of the worksheet
 * \param id : An integer representing the internal id of the
 *             sheet which is used by .xlsx revision part.
 *             (Note: id is not the index of the sheet in workbook)
 */
Worksheet::Worksheet(const QString &name, int id, Workbook *workbook) :
    d_ptr(new WorksheetPrivate(this))
{
    d_ptr->name = name;
    d_ptr->id = id;
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

void Worksheet::setHidden(bool hidden)
{
    Q_D(Worksheet);
    d->hidden = hidden;
}

int Worksheet::sheetId() const
{
    Q_D(const Worksheet);
    return d->id;
}

/*!
 * Returns whether sheet is protected.
 */
bool Worksheet::isWindowProtected() const
{
    Q_D(const Worksheet);
    return d->windowProtection;
}

/*!
 * Protects/unprotects the sheet.
 */
void Worksheet::setWindowProtected(bool protect)
{
    Q_D(Worksheet);
    d->windowProtection = protect;
}

/*!
 * Return whether formulas instead of their calculated results shown in cells
 */
bool Worksheet::isFormulasVisible() const
{
    Q_D(const Worksheet);
    return d->showFormulas;
}

/*!
 * Show formulas in cells instead of their calculated results
 */
void Worksheet::setFormulasVisible(bool visible)
{
    Q_D(Worksheet);
    d->showFormulas = visible;
}

/*!
 * Return whether gridlines is shown or not.
 */
bool Worksheet::isGridLinesVisible() const
{
    Q_D(const Worksheet);
    return d->showGridLines;
}

/*!
 * Show or hide the gridline based on \a visible
 */
void Worksheet::setGridLinesVisible(bool visible)
{
    Q_D(Worksheet);
    d->showGridLines = visible;
}

/*!
 * Return whether is row and column headers is vislbe.
 */
bool Worksheet::isRowColumnHeadersVisible() const
{
    Q_D(const Worksheet);
    return d->showRowColHeaders;
}

/*!
 * Show or hide the row column headers based on \a visible
 */
void Worksheet::setRowColumnHeadersVisible(bool visible)
{
    Q_D(Worksheet);
    d->showRowColHeaders = visible;
}


/*!
 * Return whether the sheet is shown right-to-left or not.
 */
bool Worksheet::isRightToLeft() const
{
    Q_D(const Worksheet);
    return d->rightToLeft;
}

/*!
 * Enable or disable the right-to-left.
 */
void Worksheet::setRightToLeft(bool enable)
{
    Q_D(Worksheet);
    d->rightToLeft = enable;
}

/*!
 * Return whether is cells that have zero value show a zero.
 */
bool Worksheet::isZerosVisible() const
{
    Q_D(const Worksheet);
    return d->showZeros;
}

/*!
 * Show a zero in cells that have zero value.
 */
void Worksheet::setZerosVisible(bool visible)
{
    Q_D(Worksheet);
    d->showZeros = visible;
}

/*!
 * Return whether this tab is selected.
 */
bool Worksheet::isSelected() const
{
    Q_D(const Worksheet);
    return d->tabSelected;
}

/*!
 * Set
 */
void Worksheet::setSelected(bool select)
{
    Q_D(Worksheet);
    d->tabSelected = select;
}

/*!
 * Return whether is ruler is shown.
 */
bool Worksheet::isRulerVisible() const
{
    Q_D(const Worksheet);
    return d->showRuler;

}

/*!
 * Show or hide the ruler.
 */
void Worksheet::setRulerVisible(bool visible)
{
    Q_D(Worksheet);
    d->showRuler = visible;

}

/*!
 * Return whether is outline symbols is shown.
 */
bool Worksheet::isOutlineSymbolsVisible() const
{
    Q_D(const Worksheet);
    return d->showOutlineSymbols;
}

/*!
 * Show or hide the outline symbols.
 */
void Worksheet::setOutlineSymbolsVisible(bool visible)
{
    Q_D(Worksheet);
    d->showOutlineSymbols = visible;
}

/*!
 * Return whether is white space is shown.
 */
bool Worksheet::isWhiteSpaceVisible() const
{
    Q_D(const Worksheet);
    return d->showWhiteSpace;
}

/*!
 * Show or hide the white space.
 */
void Worksheet::setWhiteSpaceVisible(bool visible)
{
    Q_D(Worksheet);
    d->showWhiteSpace = visible;
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
    } else if (value.userType() == QMetaType::Bool) { //Bool
        ret = writeBool(row,column, value.toBool(), format);
    } else if (value.toDateTime().isValid()) { //DateTime
        ret = writeDateTime(row, column, value.toDateTime(), format);
    } else if (value.toDouble(&ok), ok) { //Number
        if (!d->workbook->isStringsToNumbersEnabled() && value.userType() == QMetaType::QString) {
            //Don't convert string to number if the flag not enabled.
            ret = writeString(row, column, value.toString(), format);
        } else {
            ret = writeNumeric(row, column, value.toDouble(), format);
        }
    } else if (value.userType() == QMetaType::QUrl) { //url
        ret = writeHyperlink(row, column, value.toUrl(), format);
    } else if (value.userType() == QMetaType::QString) { //string
        QString token = value.toString();
        QRegularExpression urlPattern(QStringLiteral("^([fh]tt?ps?://)|(mailto:)|(file://)"));
        if (token.startsWith(QLatin1String("="))) {
            ret = writeFormula(row, column, token, format);
        } else if (token.contains(urlPattern)) {
            ret = writeHyperlink(row, column, QUrl(token));
        } else {
            ret = writeString(row, column, token, format);
        }
    } else { //Wrong type

        return -1;
    }

    return ret;
}

//convert the "A1" notation to row/column notation
int Worksheet::write(const QString &row_column, const QVariant &value, Format *format)
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

Format *WorksheetPrivate::cellFormat(int row, int col) const
{
    if (!cellTable.contains(row))
        return 0;
    if (!cellTable[row].contains(col))
        return 0;
    return cellTable[row][col]->format();
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

    d->sharedStrings()->addSharedString(content);
    format = format ? format : d->cellFormat(row, column);
    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(content, Cell::String, format, this));
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

    format = format ? format : d->cellFormat(row, column);
    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::InlineString, format, this));
    d->workbook->styles()->addFormat(format);
    return error;
}

int Worksheet::writeNumeric(int row, int column, double value, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    format = format ? format : d->cellFormat(row, column);
    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::Numeric, format, this));
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

    format = format ? format : d->cellFormat(row, column);
    Cell *data = new Cell(result, Cell::Formula, format, this);
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

    format = format ? format : d->cellFormat(row, column);
    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(QVariant(), Cell::Blank, format, this));
    d->workbook->styles()->addFormat(format);

    return 0;
}

int Worksheet::writeBool(int row, int column, bool value, Format *format)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    format = format ? format : d->cellFormat(row, column);
    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::Boolean, format, this));
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

    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(value, Cell::Numeric, format, this));
    d->workbook->styles()->addFormat(format);

    return 0;
}

int Worksheet::writeHyperlink(int row, int column, const QUrl &url, Format *format, const QString &display, const QString &tip)
{
    Q_D(Worksheet);
    if (d->checkDimensions(row, column))
        return -1;

    int error = 0;

    QString urlString = url.toString();

    //Generate proper display string
    QString displayString = display.isEmpty() ? urlString : display;
    if (displayString.startsWith(QLatin1String("mailto:")))
        displayString.replace(QLatin1String("mailto:"), QString());
    if (displayString.size() > d->xls_strmax) {
        displayString = displayString.left(d->xls_strmax);
        error = -2;
    }

    /*
      Location within target. If target is a workbook (or this workbook)
      this shall refer to a sheet and cell or a defined name. Can also
      be an HTML anchor if target is HTML file.

      c:\temp\file.xlsx#Sheet!A1
      http://a.com/aaa.html#aaaaa
    */
    QString locationString;
    if (url.hasFragment()) {
        locationString = url.fragment();
        urlString = url.toString(QUrl::RemoveFragment);
    }

    //Given a default style for hyperlink
    if (!format) {
        format = d->workbook->createFormat();
        format->setFontColor(Qt::blue);
        format->setFontUnderline(Format::FontUnderlineSingle);
    }

    //Write the hyperlink string as normal string.
    d->sharedStrings()->addSharedString(displayString);
    format = format ? format : d->cellFormat(row, column);
    d->cellTable[row][column] = QSharedPointer<Cell>(new Cell(displayString, Cell::String, format, this));
    d->workbook->styles()->addFormat(format);

    //Store the hyperlink data in a separate table
    d->urlTable[row][column] = new XlsxUrlData(XlsxUrlData::External, urlString, locationString, tip);

    return error;
}

bool Worksheet::addDataValidation(const DataValidation &validation)
{
    Q_D(Worksheet);
    if (validation.ranges().isEmpty() || validation.validationType()==DataValidation::None)
        return false;

    d->dataValidationsList.append(validation);
    return true;
}

int Worksheet::insertImage(int row, int column, const QImage &image, const QPointF &offset, double xScale, double yScale)
{
    Q_D(Worksheet);

    d->imageList.append(new XlsxImageData(row, column, image, offset, xScale, yScale));
    return 0;
}

/*!
    Merge a \a range of cells. The first cell should contain the data and the others should
    be blank. All cells will be applied the same style if a valid \a format is given.

    \note All cells except the top-left one will be cleared.
 */
int Worksheet::mergeCells(const CellRange &range, Format *format)
{
    Q_D(Worksheet);
    if (range.rowCount() < 2 && range.columnCount() < 2)
        return -1;

    if (d->checkDimensions(range.firstRow(), range.firstColumn()))
        return -1;

    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {
            if (row == range.firstRow() && col == range.firstColumn()) {
                Cell *cell = cellAt(row, col);
                if (cell) {
                    if (format)
                        cell->d_ptr->format = format;
                } else {
                    writeBlank(row, col, format);
                }
            } else {
                writeBlank(row, col, format);
            }
        }
    }

    if (format)
        d->workbook->styles()->addFormat(format);

    d->merges.append(range);
    return 0;
}

/*!
    \overload
    Merge a \a range of cells. The first cell should contain the data and the others should
    be blank. All cells will be applied the same style if a valid \a format is given.

    \note All cells except the top-left one will be cleared.
 */
int Worksheet::mergeCells(const QString &range, Format *format)
{
    QStringList cells = range.split(QLatin1Char(':'));
    if (cells.size() != 2)
        return -1;
    QPoint cell1 = xl_cell_to_rowcol(cells[0]);
    QPoint cell2 = xl_cell_to_rowcol(cells[1]);

    if (cell1 == QPoint(-1,-1) || cell2 == QPoint(-1, -1))
        return -1;

    return mergeCells(CellRange(cell1.x(), cell1.y(), cell2.x(), cell2.y()), format);
}

/*!
    Unmerge the cells in the \a range.
*/
int Worksheet::unmergeCells(const CellRange &range)
{
    Q_D(Worksheet);
    if (!d->merges.contains(range))
        return -1;

    d->merges.removeOne(range);
    return 0;
}

/*!
    \overload
    Unmerge the cells in the \a range.
*/
int Worksheet::unmergeCells(const QString &range)
{
    QStringList cells = range.split(QLatin1Char(':'));
    if (cells.size() != 2)
        return -1;
    QPoint cell1 = xl_cell_to_rowcol(cells[0]);
    QPoint cell2 = xl_cell_to_rowcol(cells[1]);

    if (cell1 == QPoint(-1,-1) || cell2 == QPoint(-1, -1))
        return -1;

    return unmergeCells(CellRange(cell1.x(), cell1.y(), cell2.x(), cell2.y()));
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
    if (d->windowProtection)
        writer.writeAttribute(QStringLiteral("windowProtection"), QStringLiteral("1"));
    if (d->showFormulas)
        writer.writeAttribute(QStringLiteral("showFormulas"), QStringLiteral("1"));
    if (!d->showGridLines)
        writer.writeAttribute(QStringLiteral("showGridLines"), QStringLiteral("0"));
    if (!d->showRowColHeaders)
        writer.writeAttribute(QStringLiteral("showRowColHeaders"), QStringLiteral("0"));
    if (!d->showZeros)
        writer.writeAttribute(QStringLiteral("showZeros"), QStringLiteral("0"));
    if (d->rightToLeft)
        writer.writeAttribute(QStringLiteral("rightToLeft"), QStringLiteral("1"));
    if (d->tabSelected)
        writer.writeAttribute(QStringLiteral("tabSelected"), QStringLiteral("1"));
    if (!d->showRuler)
        writer.writeAttribute(QStringLiteral("showRuler"), QStringLiteral("0"));
    if (!d->showOutlineSymbols)
        writer.writeAttribute(QStringLiteral("showOutlineSymbols"), QStringLiteral("0"));
    if (!d->showWhiteSpace)
        writer.writeAttribute(QStringLiteral("showWhiteSpace"), QStringLiteral("0"));
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
            writer.writeAttribute(QStringLiteral("min"), QString::number(col_info->firstColumn + 1));
            writer.writeAttribute(QStringLiteral("max"), QString::number(col_info->lastColumn + 1));
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
    if (d->dimension.isValid())
        d->writeSheetData(writer);
    writer.writeEndElement();//sheetData

    d->writeMergeCells(writer);
    d->writeDataValidation(writer);
    d->writeHyperlinks(writer);
    d->writeDrawings(writer);

    writer.writeEndElement();//worksheet
    writer.writeEndDocument();
}

void WorksheetPrivate::writeSheetData(XmlStreamWriter &writer)
{
    calculateSpans();
    for (int row_num = dimension.firstRow(); row_num <= dimension.lastRow(); row_num++) {
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

            for (int col_num = dimension.firstColumn(); col_num <= dimension.lastColumn(); col_num++) {
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
        int sst_idx = sharedStrings()->getSharedStringIndex(cell->value().toString());
        writer.writeAttribute(QStringLiteral("t"), QStringLiteral("s"));
        writer.writeTextElement(QStringLiteral("v"), QString::number(sst_idx));
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

    foreach (CellRange range, merges) {
        QString cell1 = xl_rowcol_to_cell(range.firstRow(), range.firstColumn());
        QString cell2 = xl_rowcol_to_cell(range.lastRow(), range.lastColumn());
        writer.writeEmptyElement(QStringLiteral("mergeCell"));
        writer.writeAttribute(QStringLiteral("ref"), cell1+QLatin1Char(':')+cell2);
    }

    writer.writeEndElement(); //mergeCells
}

void WorksheetPrivate::writeDataValidation(XmlStreamWriter &writer)
{
    if (dataValidationsList.isEmpty())
        return;

    static QMap<DataValidation::ValidationType, QString> typeMap;
    static QMap<DataValidation::ValidationOperator, QString> opMap;
    static QMap<DataValidation::ErrorStyle, QString> esMap;
    if (typeMap.isEmpty()) {
        typeMap.insert(DataValidation::None, QStringLiteral("none"));
        typeMap.insert(DataValidation::Whole, QStringLiteral("whole"));
        typeMap.insert(DataValidation::Decimal, QStringLiteral("decimal"));
        typeMap.insert(DataValidation::List, QStringLiteral("list"));
        typeMap.insert(DataValidation::Date, QStringLiteral("date"));
        typeMap.insert(DataValidation::Time, QStringLiteral("time"));
        typeMap.insert(DataValidation::TextLength, QStringLiteral("textLength"));
        typeMap.insert(DataValidation::Custom, QStringLiteral("custom"));

        opMap.insert(DataValidation::Between, QStringLiteral("between"));
        opMap.insert(DataValidation::NotBetween, QStringLiteral("notBetween"));
        opMap.insert(DataValidation::Equal, QStringLiteral("equal"));
        opMap.insert(DataValidation::NotEqual, QStringLiteral("notEqual"));
        opMap.insert(DataValidation::LessThan, QStringLiteral("lessThan"));
        opMap.insert(DataValidation::LessThanOrEqual, QStringLiteral("lessThanOrEqual"));
        opMap.insert(DataValidation::GreaterThan, QStringLiteral("greaterThan"));
        opMap.insert(DataValidation::GreaterThanOrEqual, QStringLiteral("greaterThanOrEqual"));

        esMap.insert(DataValidation::Stop, QStringLiteral("stop"));
        esMap.insert(DataValidation::Warning, QStringLiteral("warning"));
        esMap.insert(DataValidation::Information, QStringLiteral("information"));
    }

    writer.writeStartElement(QStringLiteral("dataValidations"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(dataValidationsList.size()));

    foreach (DataValidation validation, dataValidationsList) {
        writer.writeStartElement(QStringLiteral("dataValidation"));
        if (validation.validationType() != DataValidation::None)
            writer.writeAttribute(QStringLiteral("type"), typeMap[validation.validationType()]);
        if (validation.errorStyle() != DataValidation::Stop)
            writer.writeAttribute(QStringLiteral("errorStyle"), esMap[validation.errorStyle()]);
        if (validation.validationOperator() != DataValidation::Between)
            writer.writeAttribute(QStringLiteral("operator"), opMap[validation.validationOperator()]);
        if (validation.allowBlank())
            writer.writeAttribute(QStringLiteral("allowBlank"), QStringLiteral("1"));
//        if (validation.dropDownVisible())
//            writer.writeAttribute(QStringLiteral("showDropDown"), QStringLiteral("1"));
        if (validation.isPromptMessageVisible())
            writer.writeAttribute(QStringLiteral("showInputMessage"), QStringLiteral("1"));
        if (validation.isErrorMessageVisible())
            writer.writeAttribute(QStringLiteral("showErrorMessage"), QStringLiteral("1"));
        if (!validation.errorMessageTitle().isEmpty())
            writer.writeAttribute(QStringLiteral("errorTitle"), validation.errorMessageTitle());
        if (!validation.errorMessage().isEmpty())
            writer.writeAttribute(QStringLiteral("error"), validation.errorMessage());
        if (!validation.promptMessageTitle().isEmpty())
            writer.writeAttribute(QStringLiteral("promptTitle"), validation.promptMessageTitle());
        if (!validation.promptMessage().isEmpty())
            writer.writeAttribute(QStringLiteral("prompt"), validation.promptMessage());

        QStringList sqref;
        foreach (CellRange range, validation.ranges())
            sqref.append(range.toString());
        writer.writeAttribute(QStringLiteral("sqref"), sqref.join(QLatin1Char(' ')));

        if (!validation.formula1().isEmpty())
            writer.writeTextElement(QStringLiteral("formula1"), validation.formula1());
        if (!validation.formula2().isEmpty())
            writer.writeTextElement(QStringLiteral("formula2"), validation.formula2());

        writer.writeEndElement(); //dataValidation
    }

    writer.writeEndElement(); //dataValidations
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
            if (data->linkType == XlsxUrlData::External) {
                rel_count += 1;
                externUrlList.append(data->url);
                writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(rel_count));
                if (!data->location.isEmpty())
                    writer.writeAttribute(QStringLiteral("location"), data->location);
                if (!data->display.isEmpty())
                    writer.writeAttribute(QStringLiteral("display"), data->url);
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

/*!
  Sets row height and format. Row height measured in point size. If format
  equals 0 then format is ignored. \a row is zero-indexed.
 */
bool Worksheet::setRow(int row, double height, Format *format, bool hidden)
{
    Q_D(Worksheet);
    int min_col = d->dimension.firstColumn() < 0 ? 0 : d->dimension.firstColumn();

    if (d->checkDimensions(row, min_col))
        return false;

    d->rowsInfo[row] = QSharedPointer<XlsxRowInfo>(new XlsxRowInfo(height, format, hidden));
    d->workbook->styles()->addFormat(format);
    return true;
}

/*!
  \overload
  Sets row height and format. Row height measured in point size. If format
  equals 0 then format is ignored. \a row should be "1", "2", "3", ...
 */
bool Worksheet::setRow(const QString &row, double height, Format *format, bool hidden)
{
    bool ok=true;
    int r = row.toInt(&ok);
    if (ok)
        return setRow(r-1, height, format, hidden);

    return false;
}

/*!
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

    if (colFirst > colLast)
        return false;

    if (d->checkDimensions(0, colLast, ignore_row, ignore_col))
        return false;
    if (d->checkDimensions(0, colFirst, ignore_row, ignore_col))
        return false;

    QSharedPointer<XlsxColumnInfo> info(new XlsxColumnInfo(colFirst, colLast, width, format, hidden));
    d->colsInfo.append(info);

    for (int col=colFirst; col<=colLast; ++col)
        d->colsInfoHelper[col] = info;

    d->workbook->styles()->addFormat(format);

    return true;
}

/*!
  Sets column width and format for all columns from colFirst to colLast. Column
  width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. If format
  equals 0 then format is ignored. \a colFirst and \a colLast should be "A", "B", "C", ...
 */
bool Worksheet::setColumn(const QString &colFirst, const QString &colLast, double width, Format *format, bool hidden)
{
    int col1 = xl_col_name_to_value(colFirst);
    int col2 = xl_col_name_to_value(colLast);

    if (col1 == -1 || col2 == -1)
        return false;

    return setColumn(col1, col2, width, format, hidden);
}

/*!
    Return the range that contains cell data.
 */
CellRange Worksheet::dimension() const
{
    Q_D(const Worksheet);
    return d->dimension;
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

QSharedPointer<Cell> WorksheetPrivate::readNumericCellData(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("c"));

    QString v_str;
    QString f_str;
    while (!(reader.name() == QLatin1String("c") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("v"))
                v_str = reader.readElementText();
            else if (reader.name() == QLatin1String("f"))
                f_str = reader.readElementText();
        }
    }

    if (v_str.isEmpty() && f_str.isEmpty()) {
        //blank type
        return QSharedPointer<Cell>(new Cell(QVariant(), Cell::Blank));
    } else if (f_str.isEmpty()) {
        //numeric type
        return QSharedPointer<Cell>(new Cell(v_str.toDouble(), Cell::Numeric));
    } else {
        //formula type
        QSharedPointer<Cell> cell(new Cell(v_str.toDouble(), Cell::Formula));
        cell->d_ptr->formula = f_str;
        return cell;
    }
}

void WorksheetPrivate::readSheetData(XmlStreamReader &reader)
{
    Q_Q(Worksheet);
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
                    if (!format)
                        qDebug()<<QStringLiteral("<c s=\"%1\">Invalid style index: ").arg(idx)<<idx;
                }

                if (attributes.hasAttribute(QLatin1String("t"))) {
                    QString type = attributes.value(QLatin1String("t")).toString();
                    if (type == QLatin1String("s")) {
                        //string type
                        while (!(reader.name() == QLatin1String("c") && reader.tokenType() == QXmlStreamReader::EndElement)) {
                            reader.readNextStartElement();
                            if (reader.name() == QLatin1String("v")) {
                                int sst_idx = reader.readElementText().toInt();
                                sharedStrings()->incRefByStringIndex(sst_idx);
                                QString value = sharedStrings()->getSharedString(sst_idx);
                                QSharedPointer<Cell> data(new Cell(value ,Cell::String, format, q));
                                cellTable[pos.x()][pos.y()] = QSharedPointer<Cell>(data);
                            }
                        }
                    } else if (type == QLatin1String("inlineStr")) {
                        //inline string type
                        while (!(reader.name() == QLatin1String("c") && reader.tokenType() == QXmlStreamReader::EndElement)) {
                            reader.readNextStartElement();
                            if (reader.tokenType() == QXmlStreamReader::StartElement) {
                                if (reader.name() == QLatin1String("t")) {
                                    QString value = reader.readElementText();
                                    QSharedPointer<Cell> data(new Cell(value, Cell::InlineString, format, q));
                                    cellTable[pos.x()][pos.y()] = data;
                                }
                            }
                        }
                    } else if (type == QLatin1String("b")) {
                        //bool type
                        reader.readNextStartElement();
                        if (reader.name() == QLatin1String("v")) {
                            QString value = reader.readElementText();
                            QSharedPointer<Cell> data(new Cell(value.toInt() ? true : false, Cell::Boolean, format, q));
                            cellTable[pos.x()][pos.y()] = data;
                        }
                    } else if (type == QLatin1String("str")) {
                        //formula type
                        QSharedPointer<Cell> data = readNumericCellData(reader);
                        data->d_ptr->format = format;
                        data->d_ptr->parent = q;
                        cellTable[pos.x()][pos.y()] = data;
                    } else if (type == QLatin1String("e")) {
                        //error type, such as #DIV/0! #NULL! #REF! etc
                        QString v_str, f_str;
                        while (!(reader.name() == QLatin1String("c") && reader.tokenType() == QXmlStreamReader::EndElement)) {
                            reader.readNextStartElement();
                            if (reader.tokenType() == QXmlStreamReader::StartElement) {
                                if (reader.name() == QLatin1String("v"))
                                    v_str = reader.readElementText();
                                else if (reader.name() == QLatin1String("f"))
                                    f_str = reader.readElementText();
                            }
                        }
                        QSharedPointer<Cell> data(new Cell(v_str, Cell::Error, format, q));
                        if (!f_str.isEmpty())
                            data->d_ptr->formula = f_str;
                        cellTable[pos.x()][pos.y()] = data;
                    } else if (type == QLatin1String("n")) {
                        QSharedPointer<Cell> data = readNumericCellData(reader);
                        data->d_ptr->format = format;
                        data->d_ptr->parent = q;
                        cellTable[pos.x()][pos.y()] = data;
                    }
                } else {
                    //default is "n"
                    QSharedPointer<Cell> data = readNumericCellData(reader);
                    data->d_ptr->format = format;
                    data->d_ptr->parent = q;
                    cellTable[pos.x()][pos.y()] = data;
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
                info->firstColumn = min - 1;
                info->lastColumn = max - 1;

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
                for (int col=min; col<=max; ++col)
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

                    merges.append(CellRange(p0.x(), p0.y(), p1.x(), p1.y()));
                }
            }
        }
    }

    if (merges.size() != count)
        qDebug("read merge cells error");
}

void WorksheetPrivate::readDataValidations(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("dataValidations"));
    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();

    while(!(reader.name() == QLatin1String("dataValidations")
            && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement
                && reader.name() == QLatin1String("dataValidation")) {
            readDataValidation(reader);
        }
    }

    if (dataValidationsList.size() != count)
        qDebug("read data validation error");
}

void WorksheetPrivate::readDataValidation(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("dataValidation"));

    static QMap<QString, DataValidation::ValidationType> typeMap;
    static QMap<QString, DataValidation::ValidationOperator> opMap;
    static QMap<QString, DataValidation::ErrorStyle> esMap;
    if (typeMap.isEmpty()) {
        typeMap.insert(QStringLiteral("none"), DataValidation::None);
        typeMap.insert(QStringLiteral("whole"), DataValidation::Whole);
        typeMap.insert(QStringLiteral("decimal"), DataValidation::Decimal);
        typeMap.insert(QStringLiteral("list"), DataValidation::List);
        typeMap.insert(QStringLiteral("date"), DataValidation::Date);
        typeMap.insert(QStringLiteral("time"), DataValidation::Time);
        typeMap.insert(QStringLiteral("textLength"), DataValidation::TextLength);
        typeMap.insert(QStringLiteral("custom"), DataValidation::Custom);

        opMap.insert(QStringLiteral("between"), DataValidation::Between);
        opMap.insert(QStringLiteral("notBetween"), DataValidation::NotBetween);
        opMap.insert(QStringLiteral("equal"), DataValidation::Equal);
        opMap.insert(QStringLiteral("notEqual"), DataValidation::NotEqual);
        opMap.insert(QStringLiteral("lessThan"), DataValidation::LessThan);
        opMap.insert(QStringLiteral("lessThanOrEqual"), DataValidation::LessThanOrEqual);
        opMap.insert(QStringLiteral("greaterThan"), DataValidation::GreaterThan);
        opMap.insert(QStringLiteral("greaterThanOrEqual"), DataValidation::GreaterThanOrEqual);

        esMap.insert(QStringLiteral("stop"), DataValidation::Stop);
        esMap.insert(QStringLiteral("warning"), DataValidation::Warning);
        esMap.insert(QStringLiteral("information"), DataValidation::Information);
    }

    DataValidation validation;
    QXmlStreamAttributes attrs = reader.attributes();

    QString sqref = attrs.value(QLatin1String("sqref")).toString();
    foreach (QString range, sqref.split(QLatin1Char(' ')))
        validation.addRange(range);

    if (attrs.hasAttribute(QLatin1String("type"))) {
        QString t = attrs.value(QLatin1String("type")).toString();
        validation.setValidationType(typeMap.contains(t) ? typeMap[t] : DataValidation::None);
    }
    if (attrs.hasAttribute(QLatin1String("errorStyle"))) {
        QString es = attrs.value(QLatin1String("errorStyle")).toString();
        validation.setErrorStyle(esMap.contains(es) ? esMap[es] : DataValidation::Stop);
    }
    if (attrs.hasAttribute(QLatin1String("operator"))) {
        QString op = attrs.value(QLatin1String("operator")).toString();
        validation.setValidationOperator(opMap.contains(op) ? opMap[op] : DataValidation::Between);
    }
    if (attrs.hasAttribute(QLatin1String("allowBlank"))) {
        validation.setAllowBlank(true);
    } else {
        validation.setAllowBlank(false);
    }
    if (attrs.hasAttribute(QLatin1String("showInputMessage"))) {
        validation.setPromptMessageVisible(true);
    } else {
        validation.setPromptMessageVisible(false);
    }
    if (attrs.hasAttribute(QLatin1String("showErrorMessage"))) {
        validation.setErrorMessageVisible(true);
    } else {
        validation.setErrorMessageVisible(false);
    }

    QString et = attrs.value(QLatin1String("errorTitle")).toString();
    QString e = attrs.value(QLatin1String("error")).toString();
    if (!e.isEmpty() || !et.isEmpty())
        validation.setErrorMessage(e, et);

    QString pt = attrs.value(QLatin1String("promptTitle")).toString();
    QString p = attrs.value(QLatin1String("prompt")).toString();
    if (!p.isEmpty() || !pt.isEmpty())
        validation.setPromptMessage(p, pt);

    //find the end
    while(!(reader.name() == QLatin1String("dataValidation") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("formula1")) {
                validation.setFormula1(reader.readElementText());
            } else if (reader.name() == QLatin1String("formula2")) {
                validation.setFormula2(reader.readElementText());
            }
        }
    }

    dataValidationsList.append(validation);
}

void WorksheetPrivate::readSheetViews(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("sheetViews"));

    while(!(reader.name() == QLatin1String("sheetViews")
            && reader.tokenType() == QXmlStreamReader::EndElement)) {
        reader.readNextStartElement();
        if (reader.tokenType() == QXmlStreamReader::StartElement && reader.name() == QLatin1String("sheetView")) {
            QXmlStreamAttributes attrs = reader.attributes();
            //default false
            windowProtection = attrs.value(QLatin1String("windowProtection")) == QLatin1String("1");
            showFormulas = attrs.value(QLatin1String("showFormulas")) == QLatin1String("1");
            rightToLeft = attrs.value(QLatin1String("rightToLeft")) == QLatin1String("1");
            tabSelected = attrs.value(QLatin1String("tabSelected")) == QLatin1String("1");
            //default true
            showGridLines = attrs.value(QLatin1String("showGridLines")) != QLatin1String("0");
            showRowColHeaders = attrs.value(QLatin1String("showRowColHeaders")) != QLatin1String("0");
            showZeros = attrs.value(QLatin1String("showZeros")) != QLatin1String("0");
            showRuler = attrs.value(QLatin1String("showRuler")) != QLatin1String("0");
            showOutlineSymbols = attrs.value(QLatin1String("showOutlineSymbols")) != QLatin1String("0");
            showWhiteSpace = attrs.value(QLatin1String("showWhiteSpace")) != QLatin1String("0");
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
                QString range = attributes.value(QLatin1String("ref")).toString();
                d->dimension = CellRange(range);
            } else if (reader.name() == QLatin1String("sheetViews")) {
                d->readSheetViews(reader);
            } else if (reader.name() == QLatin1String("sheetFormatPr")) {

            } else if (reader.name() == QLatin1String("cols")) {
                d->readColumnsInfo(reader);
            } else if (reader.name() == QLatin1String("sheetData")) {
                d->readSheetData(reader);
            } else if (reader.name() == QLatin1String("mergeCells")) {
                d->readMergeCells(reader);
            } else if (reader.name() == QLatin1String("dataValidations")) {
                d->readDataValidations(reader);
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

/*!
 * \internal
 *  Unit test can use this member to get sharedString object.
 */
SharedStrings *WorksheetPrivate::sharedStrings() const
{
    return workbook->sharedStrings();
}

/*!
 * Return the workbook
 */
Workbook *Worksheet::workbook() const
{
    Q_D(const Worksheet);
    return d->workbook;
}

QT_END_NAMESPACE_XLSX
