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

#include "xlsxdocument.h"
#include "xlsxdocument_p.h"
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"
#include "xlsxpackage_p.h"

#include <QFile>
#include <QPointF>

QT_BEGIN_NAMESPACE_XLSX

DocumentPrivate::DocumentPrivate(Document *p) :
    q_ptr(p), defaultPackageName(QStringLiteral("Book1.xlsx"))
{
    workbook = QSharedPointer<Workbook>(new Workbook);
}

void DocumentPrivate::init()
{
    if (workbook->worksheets().size() == 0)
        workbook->addWorksheet();
}

bool DocumentPrivate::loadPackage(QIODevice *device)
{
    Q_Q(Document);
    Package package(q);
    return package.parsePackage(device);
}


/*!
  \class Document
  \inmodule QtXlsx
  \brief The Document class provides a API that is used to handle the contents of .xlsx files.

*/

/*!
 * Creates a new empty xlsx document.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    d_ptr->init();
}

/*!
 * \overload
 * Try to open an existing xlsx document named \a name.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(const QString &name, QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    d_ptr->packageName = name;
    if (QFile::exists(name)) {
        QFile xlsx(name);
        if (xlsx.open(QFile::ReadOnly))
            d_ptr->loadPackage(&xlsx);
    }
    d_ptr->init();
}

/*!
 * \overload
 * Try to open an existing xlsx document from \a device.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(QIODevice *device, QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    if (device && device->isReadable())
        d_ptr->loadPackage(device);
    d_ptr->init();
}

/*!
    \overload

    Write \a value to cell \a row_column with the \a format.
 */
int Document::write(const QString &row_column, const QVariant &value, const Format &format)
{
    return currentWorksheet()->write(row_column, value, format);
}

/*!
 * Write \a value to cell (\a row, \a col) with the \a format.
 */
int Document::write(int row, int col, const QVariant &value, const Format &format)
{
    return currentWorksheet()->write(row, col, value, format);
}

/*!
    \overload
    Returns the contents of the cell \a cell.
*/
QVariant Document::read(const QString &cell) const
{
    return currentWorksheet()->read(cell);
}

/*!
    Returns the contents of the cell (\a row, \a col).
 */
QVariant Document::read(int row, int col) const
{
    return currentWorksheet()->read(row, col);
}

/*!
 * \brief Insert an \a image to current active worksheet to the position \a row, \a column with the given
 * \a xOffset, \a yOffset, \a xScale and \a yScale.
 */
int Document::insertImage(int row, int column, const QImage &image, double xOffset, double yOffset, double xScale, double yScale)
{
    return currentWorksheet()->insertImage(row, column, image, QPointF(xOffset, yOffset), xScale, yScale);
}

/*!
    Merge a \a range of cells. The first cell should contain the data and the others should
    be blank. All cells will be applied the same style if a valid \a format is given.

    \note All cells except the top-left one will be cleared.
 */
int Document::mergeCells(const CellRange &range, const Format &format)
{
    return currentWorksheet()->mergeCells(range, format);
}

/*!
    \overload
    Merge a \a range of cells. The first cell should contain the data and the others should
    be blank. All cells will be applied the same style if a valid \a format is given.

    \note All cells except the top-left one will be cleared.
 */
int Document::mergeCells(const QString &range, const Format &format)
{
    return currentWorksheet()->mergeCells(range, format);
}

/*!
    Unmerge the cells in the \a range.
*/
int Document::unmergeCells(const QString &range)
{
    return currentWorksheet()->unmergeCells(range);
}

/*!
    Unmerge the cells in the \a range.
*/
int Document::unmergeCells(const CellRange &range)
{
    return currentWorksheet()->unmergeCells(range);
}

/*!
  Sets the properties of \a row with the given \a height, \a format and \a hidden.
  \a row is 1-indexed.

  Returns false if failed.
 */
bool Document::setRow(int row, double height, const Format &format, bool hidden)
{
    return currentWorksheet()->setRow(row, height, format, hidden);
}

/*!
  Sets the column properties for all columns from \a colFirst to \a colLast with
  the given \a width, \a format and \a hidden. Column
  width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font.
  \a colFirst and \a colLast are all 1-indexed.

  Return false if failed.
 */
bool Document::setColumn(int colFirst, int colLast, double width, const Format &format, bool hidden)
{
    return currentWorksheet()->setColumn(colFirst, colLast, width, format, hidden);
}

/*!
  \overload

  Sets column width and format for all columns from \a colFirst to \a colLast with
  the given \a width and \a format. Column
  \a width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font.
  \a colFirst and \a colLast should be "A", "B", "C", ...
 */
bool Document::setColumn(const QString &colFirst, const QString &colLast, double width, const Format &format, bool hidden)
{
    return currentWorksheet()->setColumn(colFirst, colLast, width, format, hidden);
}

/*!
   Groups rows from \a rowFirst to \a rowLast with the given \a collapsed.
   Returns false if error occurs.
 */
bool Document::groupRows(int rowFirst, int rowLast, bool collapsed)
{
    return currentWorksheet()->groupRows(rowFirst, rowLast, collapsed);
}

/*!
   Groups columns from \a colFirst to \a colLast with the given \a collapsed.
   Returns false if error occurs.
 */
bool Document::groupColumns(int colFirst, int colLast, bool collapsed)
{
    return currentWorksheet()->groupColumns(colFirst, colLast, collapsed);
}

/*!
 *  Add a data \a validation rule for current worksheet. Returns true if successful.
 */
bool Document::addDataValidation(const DataValidation &validation)
{
    return currentWorksheet()->addDataValidation(validation);
}

/*!
 * Returns a Cell object based on the given \a pos. 0 will be returned if the cell doesn't exist.
 */
Cell *Document::cellAt(const QString &pos) const
{
    return currentWorksheet()->cellAt(pos);
}

/*!
 * Returns a Cell object based on the given \a row and \a col.
 */
Cell *Document::cellAt(int row, int col) const
{
    return currentWorksheet()->cellAt(row, col);
}

/*!
 * \brief Create a defined name in the workbook with the given \a name, \a formula, \a comment
 *  and \a scope.
 *
 * \param name The defined name.
 * \param formula The cell or range that the defined name refers to.
 * \param scope The name of one worksheet, or empty which means golbal scope.
 * \return Return false if the name invalid.
 */
bool Document::defineName(const QString &name, const QString &formula, const QString &comment, const QString &scope)
{
    Q_D(Document);

    return d->workbook->defineName(name, formula, comment, scope);
}

/*!
    Return the range that contains cell data.
 */
CellRange Document::dimension() const
{
    return currentWorksheet()->dimension();
}

/*!
 * Returns the value of the document's \a key property.
 */
QString Document::documentProperty(const QString &key) const
{
    Q_D(const Document);
    if (d->documentProperties.contains(key))
        return d->documentProperties[key];
    else
        return QString();
}

/*!
    Set the document properties such as Title, Author etc.

    The method can be used to set the document properties of the Excel
    file created by Qt Xlsx. These properties are visible when you use the
    Office Button -> Prepare -> Properties option in Excel and are also
    available to external applications that read or index windows files.

    The \a property \a key that can be set are:

    \list
    \li title
    \li subject
    \li creator
    \li manager
    \li company
    \li category
    \li keywords
    \li description
    \li status
    \endlist
*/
void Document::setDocumentProperty(const QString &key, const QString &property)
{
    Q_D(Document);
    d->documentProperties[key] = property;
}

/*!
 * Returns the names of all properties that were addedusing setDocumentProperty().
 */
QStringList Document::documentPropertyNames() const
{
    Q_D(const Document);
    return d->documentProperties.keys();
}

/*!
 * Return the internal Workbook object.
 */
Workbook *Document::workbook() const
{
    Q_D(const Document);
    return d->workbook.data();
}

/*!
 * Creates and append an document with name \a name.
 * Return true if success.
 */
bool Document::addWorksheet(const QString &name)
{
    Q_D(Document);
    return d->workbook->addWorksheet(name);
}

/*!
 * Creates and inserts an document with name \a name at the \a index.
 * Returns false if the \a name already used.
 */
bool Document::insertWorkSheet(int index, const QString &name)
{
    Q_D(Document);
    return d->workbook->insertWorkSheet(index, name);
}

/*!
   Rename current worksheet to new \a name.
   Returns true if the name defined successful.
 */
bool Document::setSheetName(const QString &name)
{
    Q_D(Document);
    for (int i=0; i<d->workbook->worksheets().size(); ++i) {
        if (d->workbook->worksheets()[i]->sheetName() == name)
            return false;
    }
    currentWorksheet()->setSheetName(name);
    return true;
}

/*!
 * \brief Return pointer of current worksheet.
 */
Worksheet *Document::currentWorksheet() const
{
    Q_D(const Document);
    if (d->workbook->worksheets().size() == 0)
        return 0;

    return d->workbook->worksheets().at(d->workbook->activeWorksheet()).data();
}

/*!
 * \brief Set current worksheet to be the sheet at \a index.
 */
void Document::setCurrentWorksheet(int index)
{
    Q_D(Document);
    d->workbook->setActiveWorksheet(index);
}

/*!
 * \brief Set current worksheet to be the sheet named \a name.
 */
void Document::setCurrentWorksheet(const QString &name)
{
    Q_D(Document);
    for (int i=0; i<d->workbook->worksheets().size(); ++i) {
        if (d->workbook->worksheets()[i]->sheetName() == name)
            d->workbook->setActiveWorksheet(i);
    }
}

/*!
 * Save current document to the filesystem. If no name specified when
 * the document constructed, a default name "book1.xlsx" will be used.
 * Returns true if saved successfully.
 */
bool Document::save()
{
    Q_D(Document);
    QString name = d->packageName.isEmpty() ? d->defaultPackageName : d->packageName;

    return saveAs(name);
}

/*!
 * Saves the document to the file with the given \a name.
 * Returns true if saved successfully.
 */
bool Document::saveAs(const QString &name)
{
    QFile file(name);
    if (file.open(QIODevice::WriteOnly))
        return saveAs(&file);
    return false;
}

/*!
 * \overload
 * This function writes a document to the given \a device.
 */
bool Document::saveAs(QIODevice *device)
{
//    activedWorksheet()->setHidden(false);
//    activedWorksheet()->setSelected(true);

    //Create the package based on current workbook
    Package package(this);
    return package.createPackage(device);
}

/*!
 * Destroys the document and cleans up.
 */
Document::~Document()
{
    delete d_ptr;
}

QT_END_NAMESPACE_XLSX
