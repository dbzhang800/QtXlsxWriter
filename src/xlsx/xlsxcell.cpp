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
#include "xlsxcell.h"
#include "xlsxcell_p.h"
#include "xlsxformat.h"
#include "xlsxformat_p.h"
#include "xlsxutility_p.h"
#include "xlsxworksheet.h"
#include "xlsxworkbook.h"
#include <QDateTime>

QT_BEGIN_NAMESPACE_XLSX

CellPrivate::CellPrivate(Cell *p) :
    q_ptr(p)
{

}

CellPrivate::CellPrivate(const CellPrivate * const cp)
    : value(cp->value), formula(cp->formula), dataType(cp->dataType)
    , format(cp->format), range(cp->range), richString(cp->richString)
    , parent(cp->parent)
{

}

/*!
  \class Cell
  \inmodule QtXlsx
  \brief The Cell class provides a API that is used to handle the worksheet cell.

*/

/*!
  \enum Cell::DataType

  \value Blank,
  \value String,
  \value Numeric,
  \value Formula,
  \value Boolean,
  \value Error,
  \value InlineString,
  \value ArrayFormula
  */

/*!
 * \internal
 * Created by Worksheet only.
 */
Cell::Cell(const QVariant &data, DataType type, const Format &format, Worksheet *parent) :
    d_ptr(new CellPrivate(this))
{
    d_ptr->value = data;
    d_ptr->dataType = type;
    d_ptr->format = format;
    d_ptr->parent = parent;
}

/*!
 * \internal
 */
Cell::Cell(const Cell * const cell):
    d_ptr(new CellPrivate(cell->d_ptr))
{
    d_ptr->q_ptr = this;
}

/*!
 * Destroys the Cell and cleans up.
 */
Cell::~Cell()
{
    delete d_ptr;
}

/*!
 * Return the dataType of this Cell
 */
Cell::DataType Cell::dataType() const
{
    Q_D(const Cell);
    return d->dataType;
}

/*!
 * Return the data content of this Cell
 */
QVariant Cell::value() const
{
    Q_D(const Cell);
    return d->value;
}

/*!
 * Return the style used by this Cell. If no style used, 0 will be returned.
 */
Format Cell::format() const
{
    Q_D(const Cell);
    return d->format;
}

/*!
 * Return the formula contents if the dataType is Formula
 */
QString Cell::formula() const
{
    Q_D(const Cell);
    return d->formula;
}

/*!
 * Returns whether the value is probably a dateTime or not
 */
bool Cell::isDateTime() const
{
    Q_D(const Cell);
    if (d->dataType == Numeric && d->value.toDouble() >=0
            && d->format.isValid() && d->format.isDateTimeFormat()) {
        return true;
    }
    return false;
}

/*!
 * Return the data time value.
 */
QDateTime Cell::dateTime() const
{
    Q_D(const Cell);
    if (!isDateTime())
        return QDateTime();
    return datetimeFromNumber(d->value.toDouble(), d->parent->workbook()->isDate1904());
}

/*!
 * Returns whether the cell is probably a rich string or not
 */
bool Cell::isRichString() const
{
    Q_D(const Cell);
    if (d->dataType != String && d->dataType != InlineString)
        return false;

    return d->richString.isRichString();
}

QT_END_NAMESPACE_XLSX
