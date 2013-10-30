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

#include "xlsxdatavalidation.h"
#include "xlsxdatavalidation_p.h"
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"

QT_BEGIN_NAMESPACE_XLSX

DataValidationPrivate::DataValidationPrivate()
    :validationType(DataValidation::None), validationOperator(DataValidation::Between)
    , errorStyle(DataValidation::Stop), allowBlank(false), isPromptMessageVisible(true)
    , isErrorMessageVisible(true)
{

}

DataValidationPrivate::DataValidationPrivate(DataValidation::ValidationType type, DataValidation::ValidationOperator op, const QString &formula1, const QString &formula2, bool allowBlank)
    :validationType(type), validationOperator(op)
    , errorStyle(DataValidation::Stop), allowBlank(allowBlank), isPromptMessageVisible(true)
    , isErrorMessageVisible(true), formula1(formula1), formula2(formula2)
{

}

DataValidationPrivate::DataValidationPrivate(const DataValidationPrivate &other)
    :QSharedData(other)
    , validationType(DataValidation::None), validationOperator(DataValidation::Between)
    , errorStyle(DataValidation::Stop), allowBlank(false), isPromptMessageVisible(true)
    , isErrorMessageVisible(true)
{

}

DataValidationPrivate::~DataValidationPrivate()
{

}

/*!
 * \class DataValidation
 * \brief Data validation for single cell or a range
 * \inmodule QtXlsx
 * The data validation can be applied to a single cell or a range of cells.
 */

/*!
 * \enum DataValidation::ValidationType
 *
 * The enum type defines the type of data that you wish to validate.
 *
 * \value None the type of data is unrestricted. This is the same as not applying a data validation.
 * \value Whole restricts the cell to integer values. Means "Whole number"?
 * \value Decimal restricts the cell to decimal values.
 * \value List restricts the cell to a set of user specified values.
 * \value Date restricts the cell to date values.
 * \value Time restricts the cell to time values.
 * \value TextLength restricts the cell data based on an integer string length.
 * \value Custom restricts the cell based on an external Excel formula that returns a true/false value.
 */

/*!
 * \enum DataValidation::ValidationOperator
 *
 *  The enum type defines the criteria by which the data in the
 *  cell is validated
 *
 * \value Between
 * \value NotBetween
 * \value Equal
 * \value NotEqual
 * \value LessThan
 * \value LessThanOrEqual
 * \value GreaterThan
 * \value GreaterThanOrEqual
 */

/*!
 * \enum DataValidation::ErrorStyle
 *
 *  The enum type defines the type of error dialog that
 *  is displayed.
 *
 * \value Stop
 * \value Warning
 * \value Information
 */

/*!
 * Construct a data validation object
 */
DataValidation::DataValidation(ValidationType type, ValidationOperator op, const QString &formula1, const QString &formula2, bool allowBlank)
    :d(new DataValidationPrivate(type, op, formula1, formula2, allowBlank))
{

}

/*!
    Construct a data validation object
*/
DataValidation::DataValidation()
    :d(new DataValidationPrivate())
{

}

/*!
    \internal
*/
DataValidation::DataValidation(const DataValidation &other)
    :d(other.d)
{

}

/*!
 * Destroy the object.
 */
DataValidation::~DataValidation()
{
}

DataValidation::ValidationType DataValidation::validationType() const
{
    return d->validationType;
}

DataValidation::ValidationOperator DataValidation::validationOperator() const
{
    return d->validationOperator;
}

DataValidation::ErrorStyle DataValidation::errorStyle() const
{
    return d->errorStyle;
}

QString DataValidation::formula1() const
{
    return d->formula1;
}

QString DataValidation::formula2() const
{
    return d->formula2;
}

bool DataValidation::allowBlank() const
{
    return d->allowBlank;
}

QString DataValidation::errorMessage() const
{
    return d->errorMessage;
}

QString DataValidation::errorMessageTitle() const
{
    return d->errorMessageTitle;
}

QString DataValidation::promptMessage() const
{
    return d->promptMessage;
}

QString DataValidation::promptMessageTitle() const
{
    return d->promptMessageTitle;
}

bool DataValidation::isPromptMessageVisible() const
{
    return d->isPromptMessageVisible;
}

bool DataValidation::isErrorMessageVisible() const
{
    return d->isErrorMessageVisible;
}

QList<CellRange> DataValidation::ranges() const
{
    return d->ranges;
}

void DataValidation::setValidationType(DataValidation::ValidationType type)
{
    d->validationType = type;
}

void DataValidation::setValidationOperator(DataValidation::ValidationOperator op)
{
    d->validationOperator = op;
}

void DataValidation::setErrorStyle(DataValidation::ErrorStyle es)
{
    d->errorStyle = es;
}

void DataValidation::setFormula1(const QString &formula)
{
    if (formula.startsWith(QLatin1Char('=')))
        d->formula1 = formula.mid(1);
    else
        d->formula1 = formula;
}

void DataValidation::setFormula2(const QString &formula)
{
    if (formula.startsWith(QLatin1Char('=')))
        d->formula2 = formula.mid(1);
    else
        d->formula2 = formula;}

void DataValidation::setErrorMessage(const QString &error, const QString &title)
{
    d->errorMessage = error;
    d->errorMessageTitle = title;
}

void DataValidation::setPromptMessage(const QString &prompt, const QString &title)
{
    d->promptMessage = prompt;
    d->promptMessageTitle = title;
}

void DataValidation::setAllowBlank(bool enable)
{
    d->allowBlank = enable;
}

void DataValidation::setPromptMessageVisible(bool visible)
{
    d->isPromptMessageVisible = visible;
}

void DataValidation::setErrorMessageVisible(bool visible)
{
    d->isErrorMessageVisible = visible;
}

/*!
    Add the \a cell which the DataValidation will apply to.
 */
void DataValidation::addCell(const QString &cell)
{
    d->ranges.append(CellRange(cell));
}

/*!
    \overload
 */
void DataValidation::addCell(int row, int col)
{
    d->ranges.append(CellRange(row, col, row, col));
}

/*!
    Add the \a range which the DataValidation will apply to.
 */
void DataValidation::addRange(const QString &range)
{
    d->ranges.append(CellRange(range));
}

/*!
    \overload
 */
void DataValidation::addRange(int firstRow, int firstCol, int lastRow, int lastCol)
{
    d->ranges.append(CellRange(firstRow, firstCol, lastRow, lastCol));
}

/*!
    \overload
 */
void DataValidation::addRange(const CellRange &range)
{
    d->ranges.append(range);
}

QT_END_NAMESPACE_XLSX
