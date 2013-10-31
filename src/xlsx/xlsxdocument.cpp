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
 * Create a new format used by sheet cells.
 */
Format *Document::createFormat()
{
    Q_D(Document);
    return d->workbook->createFormat();
}

/*!
 * Write \a value to cell \a row_column with the \a format.
 */
int Document::write(const QString &row_column, const QVariant &value, Format *format)
{
    return currentWorksheet()->write(row_column, value, format);
}

/*!
 * Write \a value to cell (\a row, \a col) with the \a format.
 */
int Document::write(int row, int col, const QVariant &value, Format *format)
{
    return currentWorksheet()->write(row, col, value, format);
}

/*!
 * \brief Insert an image to current active worksheet.
 * \param row
 * \param column
 * \param image
 * \param xOffset
 * \param yOffset
 * \param xScale
 * \param yScale
 */
int Document::insertImage(int row, int column, const QImage &image, double xOffset, double yOffset, double xScale, double yScale)
{
    return currentWorksheet()->insertImage(row, column, image, QPointF(xOffset, yOffset), xScale, yScale);
}

/*!
 * Merge cell \a range.
 */
int Document::mergeCells(const QString &range)
{
    return currentWorksheet()->mergeCells(range);
}

/*!
 * Unmerge cell \a range.
 */
int Document::unmergeCells(const QString &range)
{
    return currentWorksheet()->unmergeCells(range);
}

/*!
 * \brief Set properties for a row of cells.
 * \param row The worksheet row (zero indexed).
 * \param height The row height.
 * \param format Optional Format object.
 * \param hidden
 */
bool Document::setRow(int row, double height, Format *format, bool hidden)
{
    return currentWorksheet()->setRow(row, height, format, hidden);
}

/*!
  \overload
  Sets row height and format. Row height measured in point size. If format
  equals 0 then format is ignored. \a row should be "1", "2", "3", ...
 */
bool Document::setRow(const QString &row, double height, Format *format, bool hidden)
{
    return currentWorksheet()->setRow(row, height, format, hidden);
}

/*!
  Sets column width and format for all columns from colFirst to colLast. Column
  width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. If format
  equals 0 then format is ignored. \a colFirst and \a colLast are all zero-indexed.
 */
bool Document::setColumn(int colFirst, int colLast, double width, Format *format, bool hidden)
{
    return currentWorksheet()->setColumn(colFirst, colLast, width, format, hidden);
}

/*!
  Sets column width and format for all columns from colFirst to colLast. Column
  width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. If format
  equals 0 then format is ignored. \a colFirst and \a colLast should be "A", "B", "C", ...
 */
bool Document::setColumn(const QString &colFirst, const QString &colLast, double width, Format *format, bool hidden)
{
    return currentWorksheet()->setColumn(colFirst, colLast, width, format, hidden);
}

/*!
 * \brief Add a data validation rule for current worksheet
 * \param validation
 * \return
 */
bool Document::addDataValidation(const DataValidation &validation)
{
    return currentWorksheet()->addDataValidation(validation);
}

/*!
 * Returns a Cell object based on the given \a pos.
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
 * \brief Create a defined name in the workbook.
 * \param name The defined name
 * \param formula The cell or range that the defined name refers to.
 * \param comment
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

    The properties \a key that can be set are:

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
 * \brief Rename current worksheet to new \a name.
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
 */
bool Document::save()
{
    Q_D(Document);
    QString name = d->packageName.isEmpty() ? d->defaultPackageName : d->packageName;

    return saveAs(name);
}

/*!
 * Saves the document to the file with the given \a name.
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
